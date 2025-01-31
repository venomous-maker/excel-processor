use actix_web::{web, App, HttpResponse, HttpServer, Error};
use actix_files::Files; // For serving static files
use calamine::{open_workbook_auto_from_rs, Reader, Sheets};
use xlsxwriter::*;
use zip::{ZipArchive, ZipWriter};
use std::fs::File;
use std::io::{Cursor, Read, Write};
use std::path::Path;
use futures::StreamExt;
use serde::Serialize;
use std::sync::Mutex;
use std::collections::HashMap;
use chrono::{Local, NaiveDateTime};
use std::fs;

#[derive(Serialize)]
struct ApiResponse {
    message: String,
}

#[derive(Serialize, Clone)]
struct FileInfo {
    name: String,
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
async fn upload_files(mut payload: web::Payload, data: web::Data<AppState>) -> Result<HttpResponse, Error> {
    let mut files = Vec::new();

    // Read the uploaded files
    while let Some(chunk) = payload.next().await {
        let chunk = chunk?;
        files.extend_from_slice(&chunk);
    }

    // Check if the uploaded file is a ZIP file based on extension
    let is_zip = match String::from_utf8(files.clone()) {
        Ok(file_content) => file_content.ends_with(".zip"), // Check if the file has a .zip extension
        Err(_) => false, // Handle invalid UTF-8 content
    };

    if is_zip {
        // Unzip the files and process each file
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

            // Check if the file is an Excel file
            if file_name.ends_with(".xlsx") || file_name.ends_with(".xls") {
                // Read the file content
                let mut file_data = Vec::new();
                file.read_to_end(&mut file_data)?;

                // Process the file
                match process_excel_files(&file_data) {
                    Ok(output_file) => {
                        // Add the processed file to the in-memory storage
                        let mut files_map = data.files.lock().unwrap();
                        let mut next_id = data.next_id.lock().unwrap();
                        files_map.insert(*next_id, FileInfo {
                            name: output_file.clone(),
                        });
                        *next_id += 1;

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

        Ok(HttpResponse::Ok().json(ApiResponse {
            message: format!("ZIP file processed successfully. Processed {} Excel files.", processed_files.len()),
        }))
    } else {
        // Process a single Excel file
        match process_excel_files(&files) {
            Ok(output_file) => {
                // Add the file to the in-memory storage
                let mut files_map = data.files.lock().unwrap();
                let mut next_id = data.next_id.lock().unwrap();
                files_map.insert(*next_id, FileInfo {
                    name: output_file.clone(),
                });
                *next_id += 1;

                Ok(HttpResponse::Ok().json(ApiResponse {
                    message: "File processed successfully".to_string(),
                }))
            }
            Err(e) => {
                eprintln!("Non zip Failed to process file: {}", e);
                Ok(HttpResponse::BadRequest().json(ApiResponse {
                    message: format!("Non zip Failed to process file: {}", e),
                }))
            }
        }
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
    let files_map = data.files.lock().unwrap();
    let files: Vec<FileInfo> = files_map.values().cloned().collect();
    Ok(HttpResponse::Ok().json(files))
}

// Process Excel files
fn process_excel_files(file_data: &[u8]) -> Result<String, Box<dyn std::error::Error>> {
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
            // Serve static files from the "static" directory
            .service(Files::new("/", "./static").index_file("index.html"))
    })
        .bind("127.0.0.1:8000")?
        .run()
        .await
}