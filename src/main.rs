use actix_web::{web, App, HttpResponse, HttpServer, Error};
use actix_files::Files; // For serving static files
use calamine::{open_workbook_auto_from_rs, Reader, Sheets, Data};
use actix_multipart::Multipart;
use futures_util::StreamExt;
use std::path::Path;
use xlsxwriter::*;
use zip::{ZipArchive, ZipWriter};
use std::fs::File;
use std::io::{Cursor, Read, Write};
// use futures::StreamExt;
use serde::{Deserialize, Serialize};
use std::sync::Mutex;
use std::collections::HashMap;
use chrono::{Local, NaiveDateTime};
use std::fs;
use rayon::prelude::*; // Import Rayon parallel iterators


#[derive(Serialize, Clone)]
struct SearchResult {
    sheet_name: String,
    row: usize,
    col: usize,
    value: String,
    file: String,
}

#[derive(Serialize)]
struct ApiResponse {
    message: String,
}

#[derive(Serialize, Clone)]
struct FileInfo {
    name: String,
}

#[derive(Deserialize, Clone)]
struct SearchQuery {
    query: String,
}


// In-memory storage for files (for demonstration purposes)
struct AppState {
    files: Mutex<HashMap<usize, FileInfo>>,
    next_id: Mutex<usize>,
}

fn output_directory(dir_path: &str) -> & str {
    // Use default directory if dir_path is empty
    let path = if dir_path.is_empty() {
        "output_files"
    } else {
        dir_path
    };

    // Create directory if it doesn't exist
    if !Path::new(&path).exists() {
        fs::create_dir_all(&path).expect("Failed to create output directory");
    }

    path
}

// Handler for uploading and processing Excel files
async fn upload_files(mut payload: Multipart, data: web::Data<AppState>) -> Result<HttpResponse, Error> {
    while let Some(field) = payload.next().await {
        let mut field = field?;
        let content_disposition = field.content_disposition();

        // Extract filename from content disposition header
        let file_name = content_disposition
            .unwrap().get_filename()
            .map(|name| name.to_string())
            .unwrap_or_else(|| "unknown_file".to_string());

        // Extract file extension
        let file_extension = Path::new(&file_name)
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or("")
            .to_string();

        let mut files = Vec::new();

        // Read the file content
        while let Some(chunk) = field.next().await {
            let chunk = chunk?;
            files.extend_from_slice(&chunk);
        }

        // Log filename and extension
        println!("Uploaded file: {} (Extension: {})", file_name, file_extension);

        // let is_zip = files.len() >= 4 && &files[0..4] == b"PK\x03\x04";
        let is_zip = file_extension == "zip";

        let mut zip_buffer = vec![];

        if is_zip {
            println!("Detected ZIP file, processing...");

            // Handle ZIP file processing
            let cursor = Cursor::new(files);
            let mut archive = ZipArchive::new(cursor).map_err(|e| {
                actix_web::error::ErrorInternalServerError(format!("Failed to open ZIP archive: {}", e))
            })?;

            let mut processed_files = Vec::new();
            for i in 0..archive.len() {
                let mut file = archive.by_index(i).map_err(|e| {
                    actix_web::error::ErrorInternalServerError(format!("Failed to read file in ZIP archive: {}", e))
                })?;
                let file_name = file.name().to_string();

                if file_name.ends_with(".xlsx") || file_name.ends_with(".xls") {
                    let mut file_data = Vec::new();
                    file.read_to_end(&mut file_data)?;

                    match process_excel_files(&file_data) {
                        Ok(output_file) => {
                            let mut files_map = data.files.lock().unwrap();
                            let mut next_id = data.next_id.lock().unwrap();
                            files_map.insert(*next_id, FileInfo {
                                name: output_file.clone(),
                            });
                            *next_id += 1;
                            zip_buffer = zip_files(&[output_file.clone()])?;
                            processed_files.push(output_file);
                        }
                        Err(e) => {
                            eprintln!("Failed to process file {}: {}", file_name, e);
                        }
                    }
                } else {
                    eprintln!("Skipping non-Excel file: {}", file_name);
                }
            }

            return Ok(HttpResponse::Ok()
                .content_type("application/zip")
                .body(zip_buffer));
        } else if file_extension == "xlsx" || file_extension == "xls" {
            println!("Detected Excel file, processing...");

            // Process non-ZIP Excel file
            match process_excel_files(&files) {
                Ok(output_file) => {
                    let mut files_map = data.files.lock().unwrap();
                    let mut next_id = data.next_id.lock().unwrap();
                    files_map.insert(*next_id, FileInfo {
                        name: output_file.clone(),
                    });
                    *next_id += 1;

                    zip_buffer = zip_files(&[output_file.clone()])?;
                    return Ok(HttpResponse::Ok()
                        .content_type("application/zip")
                        .body(zip_buffer));
                }
                Err(e) => {
                    eprintln!("Failed to process file: {}", e);
                    return Ok(HttpResponse::BadRequest().json(ApiResponse {
                        message: format!("Failed to process file: {}", e),
                    }));
                }
            }
        } else {
            return Ok(HttpResponse::BadRequest().json(ApiResponse {
                message: format!("Unsupported file type: {}", file_extension),
            }));
        }
    }

    Ok(HttpResponse::BadRequest().json(ApiResponse {
        message: "No files uploaded".to_string(),
    }))
}

// Corrected search_files function
async fn search_files(
    query: web::Query<SearchQuery>,
    data: web::Data<AppState>,
) -> Result<HttpResponse, Error> {
    let query = &query.query; // Extract the query string
    let mut files_map = data.files.lock().unwrap();
    let mut results = Vec::new();

    for file_info in files_map.values() {
        let file_path = "".to_owned() + &*file_info.name.clone(); // Simulated file path
        println!("Searching file: {}", file_path);
        // Read and process each file
        let file_data = std::fs::read(file_path).map_err(|e| {
            actix_web::error::ErrorInternalServerError(format!("Failed to read file: {}", e))
        })?;

        let cursor = Cursor::new(file_data);

        // Explicitly handle calamine::Error and convert it to actix_web::Error
        let mut workbook = match open_workbook_auto_from_rs(cursor) {
            Ok(wb) => wb,
            Err(e) => {
                return Err(actix_web::error::ErrorInternalServerError(format!(
                    "Failed to open workbook: {}",
                    e
                )));
            }
        };

        // Iterate over each sheet and search for the query string
        for sheet_name in workbook.sheet_names().to_owned() {
            // Handle Result explicitly instead of expecting Option
            match workbook.worksheet_range(&sheet_name) {
                Ok(range) => {
                    for (row_idx, row) in range.rows().enumerate() {
                        for (col_idx, cell) in row.iter().enumerate() {
                            // Search for matching data
                            let cell_value = match cell {
                                calamine::Data::String(s) => s.clone(),
                                calamine::Data::Float(f) => f.to_string(),
                                calamine::Data::Int(i) => i.to_string(),
                                calamine::Data::Bool(b) => {
                                    if *b {
                                        "TRUE".to_string()
                                    } else {
                                        "FALSE".to_string()
                                    }
                                }
                                calamine::Data::DateTime(d) => d.to_string(),
                                _ => "".to_string(),
                            };

                            // Check if the cell value contains the query string
                            if cell_value.contains(&*query) {
                                results.push(SearchResult {
                                    sheet_name: sheet_name.clone(),
                                    row: row_idx,
                                    col: col_idx,
                                    value: cell_value.clone(),
                                    file: file_info.name.clone(),
                                });
                            }
                        }
                    }
                }
                Err(e) => {
                    // Log/skip sheets with errors (optional)
                    eprintln!("Error reading range for sheet '{}': {}", sheet_name, e);
                    continue;
                }
            }
        }
    }

    if results.is_empty() {
        Ok(HttpResponse::Ok().json("No results found".to_string()))
    } else {
        Ok(HttpResponse::Ok().json(serde_json::json!({
        "data": results,
        "count": results.len()
        })))
    }
}



// Handler for deleting a file
async fn delete_file(data: web::Data<AppState>, index: web::Path<usize>) -> Result<HttpResponse, Error> {
    let mut files_map = data.files.lock().unwrap();

    // Check if the file exists in the in-memory storage
    if let Some(file_info) = files_map.remove(&index) {
        // Delete the file from the filesystem
        if fs::remove_file(&file_info.name).is_ok() {
            Ok(HttpResponse::Ok().json(ApiResponse {
                message: "File deleted successfully".to_string(),
            }))
        } else {
            // If file deletion fails, reinsert the file into the in-memory storage
            files_map.insert(index.into_inner(), file_info);
            Ok(HttpResponse::InternalServerError().json(ApiResponse {
                message: "Failed to delete file from the filesystem".to_string(),
            }))
        }
    } else {
        Ok(HttpResponse::NotFound().json(ApiResponse {
            message: "File not found".to_string(),
        }))
    }
}

// Handler for fetching the list of files
async fn get_files(data: web::Data<AppState>) -> Result<HttpResponse, Error> {
    let mut files_map = data.files.lock().unwrap();
    let mut file_list: Vec<FileInfo> = files_map.values().cloned().collect();

    let output_dir = output_directory("output_files");
    if let Ok(entries) = fs::read_dir(output_dir) {
        for entry in entries.flatten() {
            if let Ok(metadata) = entry.metadata() {
                if metadata.is_file() {
                    let file_name = output_dir.to_string()+"/"+ &*entry.file_name().into_string().unwrap();
                    let exists = file_list.iter().any(|f| f.name == file_name);
                    if !exists {
                        file_list.push(FileInfo { name: file_name.clone() });

                        // Add to in-memory storage for consistency
                        let mut next_id = data.next_id.lock().unwrap();
                        files_map.insert(*next_id, FileInfo { name: file_name });
                        *next_id += 1;
                    }
                }
            }
        }
    }

    Ok(HttpResponse::Ok().json(file_list))
}

fn process_excel_files(file_data: &[u8]) -> Result<String, Box<dyn std::error::Error>> {
    let cursor = Cursor::new(file_data);

    // Use `open_workbook_auto_from_rs` to read from an in-memory buffer
    let mut workbook: calamine::Sheets<_> = calamine::open_workbook_auto_from_rs(cursor)?;

    // Read the first sheet
    let sheet_name = workbook.sheet_names()[0].clone();
    let range = workbook.worksheet_range(&sheet_name)?;

    // Create a new output Excel file
    let output_dir = output_directory("output_files");
    let output_file = format!("{}/firstsheet{}.xlsx", output_dir, Local::now().format("%m%d%y%H%M%S"));
    let workbook = Workbook::new(&output_file)?;
    let mut sheet = workbook.add_worksheet(None)?;

    // Convert rows to a Vec for parallel processing
    // Convert to vector for Rayon parallel iteration
    let rows: Vec<_> = range.rows().enumerate().collect();

    // Use parallel iteration to process the rows
    let data: Vec<(usize, usize, String)> = rows
        .into_par_iter()
        .flat_map(|(row_idx, row)| {
            row.iter()
                .enumerate()
                .filter_map(move |(col_idx, cell)| {
                    match cell {
                        Data::String(s) => Some((row_idx, col_idx, s.clone())), // Keep string as is
                        Data::Float(f) => Some((row_idx, col_idx, f.to_string())), // Convert float to string
                        Data::Int(i) => Some((row_idx, col_idx, i.to_string())), // Convert integer to string
                        Data::Bool(b) => Some((row_idx, col_idx, if *b { "TRUE".to_string() } else { "FALSE".to_string() })), // Convert bool to string
                        Data::DateTime(d) => d.as_datetime().map(|naive_dt| {
                            (row_idx, col_idx, naive_dt.format("%Y-%m-%d %H:%M:%S").to_string())
                        }), // Format datetime
                        Data::Error(e) => Some((row_idx, col_idx, format!("Error: {:?}", e))), // Handle errors
                        Data::Empty => None, // Skip empty cells
                        _ => None
                    }
                })
                .collect::<Vec<_>>() // Collect within flat_map
        })
        .collect();

    // Sequentially write the collected data
    for (row_idx, col_idx, cell) in data {
        sheet.write_string(row_idx as u32, col_idx as u16, &cell, None)?;
    }

    workbook.close()?;
    Ok(output_file)
}


// Process Excel files
fn process_excel_files_(file_data: &[u8]) -> Result<String, Box<dyn std::error::Error>> {
    let cursor = Cursor::new(file_data);

    // Use `open_workbook_auto_from_rs` to read from an in-memory buffer
    let mut workbook: Sheets<_> = open_workbook_auto_from_rs(cursor)?;

    // Read the first sheet
    let sheet_name = workbook.sheet_names()[0].clone();
    let range = workbook.worksheet_range(&sheet_name)?;

    // Create a new output Excel file
    let output_dir = output_directory("output_files");
    let output_file = format!("{}/firstsheet{}.xlsx", output_dir, Local::now().format("%m%d%y%H%M%S"));
    let workbook = Workbook::new(&output_file)?;
    let mut sheet = workbook.add_worksheet(None)?;

    // Write data to the output file
    for (row_idx, row) in range.rows().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            match cell {
                calamine::Data::String(s) => {
                    let cleaned = s.replace(" ", ""); // Remove spaces
                    sheet.write_string(row_idx as u32, col_idx as u16, &cleaned, None)?;
                }
                calamine::Data::Float(f) => {
                    sheet.write_number(row_idx as u32, col_idx as u16, *f, None)?;
                }
                calamine::Data::Int(i) => {
                    sheet.write_number(row_idx as u32, col_idx as u16, *i as f64, None)?;
                }
                calamine::Data::Bool(b) => {
                    let bool_text = if *b { "TRUE" } else { "FALSE" };
                    sheet.write_string(row_idx as u32, col_idx as u16, bool_text, None)?;
                }
                calamine::Data::DateTime(d) => {
                    if let Some(naive_dt) = d.as_datetime() {
                        let formatted_date = naive_dt.format("%Y-%m-%d %H:%M:%S").to_string();
                        sheet.write_string(row_idx as u32, col_idx as u16, &formatted_date, None)?;
                    }
                }
                calamine::Data::Error(e) => {
                    sheet.write_string(row_idx as u32, col_idx as u16, &format!("Error: {:?}", e), None)?;
                }
                calamine::Data::Empty => {
                    // Leave the cell empty
                }
                _ => {}
            }
        }
    }

    workbook.close()?;
    Ok(output_file)
}

// Zip files into a single archive
fn zip_files(file_paths: &[String]) -> Result<Vec<u8>, Box<dyn std::error::Error>> {
    let mut zip_buffer = Vec::new();
    let mut zip_writer = ZipWriter::new(Cursor::new(&mut zip_buffer));

    for file_path in file_paths {
        let file_name = Path::new(file_path).file_name().unwrap().to_str().unwrap();
        zip_writer.start_file::<_, ()>(file_name, zip::write::FileOptions::default().into())?;
        let mut file = File::open(file_path)?;
        std::io::copy(&mut file, &mut zip_writer)?;
    }

    zip_writer.finish()?;
    Ok(zip_buffer)
}

#[actix_web::main]
async fn main() -> std::io::Result<()> {
    // Initialize the shared state
    let app_state = web::Data::new(AppState {
        files: Mutex::new(HashMap::new()),
        next_id: Mutex::new(0),
    });

    // Start the Actix-web server
    HttpServer::new(move || {
        App::new()
            .app_data(app_state.clone()) // Share the state with all routes
            // API endpoint for file upload
            .route("/upload", web::post().to(upload_files))
            // API endpoint for deleting a file
            .route("/delete/{index}", web::delete().to(delete_file))
            // API endpoint for fetching the list of files
            .route("/files", web::get().to(get_files))
            .route("/search", web::get().to(search_files)) // Add the search endpoint
            // Serve static files from the "static" directory
            .service(Files::new("/", "./static").index_file("index.html"))
    })
        .bind("127.0.0.1:8000")?
        .run()
        .await
}