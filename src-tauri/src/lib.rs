// Этот файл является частью time-to-table
// SPDX-License-Identifier: GPL-3.0-or-later

// Подробнее о командах Tauri: https://tauri.app/develop/calling-rust/

use std::path::PathBuf;
use std::sync::{Mutex, LazyLock};
use std::time::{Duration, Instant};
use std::collections::HashMap;

// Rate limiting: максимум 10 операций в секунду на команду
const MAX_CALLS_PER_SECOND: usize = 10;
const RATE_LIMIT_WINDOW: Duration = Duration::from_secs(1);

// Максимальный размер файла: 10MB
const MAX_FILE_SIZE: usize = 10 * 1024 * 1024;

struct RateLimiter {
    calls: HashMap<String, Vec<Instant>>,
}

impl RateLimiter {
    fn new() -> Self {
        RateLimiter {
            calls: HashMap::new(),
        }
    }

    fn check_rate_limit(&mut self, command: &str) -> Result<(), String> {
        let now = Instant::now();
        let key = command.to_string();
        
        // Получаем или создаём список вызовов для этой команды
        let timestamps = self.calls.entry(key).or_insert_with(Vec::new);
        
        // Удаляем старые временные метки (старше 1 секунды)
        timestamps.retain(|&t| now.duration_since(t) < RATE_LIMIT_WINDOW);
        
        // Проверяем лимит
        if timestamps.len() >= MAX_CALLS_PER_SECOND {
            return Err("Превышен лимит запросов. Попробуйте позже.".to_string());
        }
        
        // Добавляем текущую временную метку
        timestamps.push(now);
        Ok(())
    }
}

static RATE_LIMITER: LazyLock<Mutex<RateLimiter>> = LazyLock::new(|| Mutex::new(RateLimiter::new()));

/// Проверяет что путь находится в разрешённой директории
fn is_path_allowed(path: &PathBuf) -> bool {
    let allowed_dirs: Vec<PathBuf> = [
        dirs::download_dir(),
        dirs::document_dir(),
        dirs::desktop_dir(),
    ]
    .into_iter()
    .flatten()
    .collect();

    // Канонизируем путь для защиты от ../ атак
    let canonical = match path.canonicalize() {
        Ok(p) => p,
        Err(_) => {
            // Если файл ещё не существует, проверяем родительскую директорию
            if let Some(parent) = path.parent() {
                match parent.canonicalize() {
                    Ok(p) => p,
                    Err(_) => return false,
                }
            } else {
                return false;
            }
        }
    };

    allowed_dirs.iter().any(|dir| {
        if let Ok(canonical_dir) = dir.canonicalize() {
            canonical.starts_with(&canonical_dir)
        } else {
            false
        }
    })
}

/// Безопасная запись файла с проверкой пути, размера и rate limiting
#[tauri::command]
fn save_file_secure(path: String, content: String) -> Result<String, String> {
    // Rate limiting
    if let Ok(mut limiter) = RATE_LIMITER.lock() {
        limiter.check_rate_limit("save_file_secure")?;
    } else {
        return Err("Ошибка доступа к rate limiter".into());
    }
    
    // Проверка размера контента
    if content.len() > MAX_FILE_SIZE {
        return Err(format!("Размер файла превышает максимальный ({} МБ)", MAX_FILE_SIZE / 1024 / 1024));
    }
    
    let path_buf = PathBuf::from(&path);
    
    // Проверка расширения файла (только .json и .xml)
    if let Some(ext) = path_buf.extension() {
        let ext_str = ext.to_string_lossy().to_lowercase();
        if ext_str != "json" && ext_str != "xml" {
            return Err("Разрешена запись только .json и .xml файлов".into());
        }
    } else {
        return Err("Файл должен иметь расширение".into());
    }
    
    if !is_path_allowed(&path_buf) {
        return Err("Сохранение разрешено только в папки: Загрузки, Документы или Рабочий стол".into());
    }
    
    std::fs::write(&path_buf, &content)
        .map_err(|e| format!("Ошибка записи: {}", e))?;
    
    Ok(path)
}

/// Безопасная запись бинарного файла (для .xlsx) с проверкой пути, размера и rate limiting
#[tauri::command]
fn save_file_binary(path: String, content: Vec<u8>) -> Result<String, String> {
    // Rate limiting
    if let Ok(mut limiter) = RATE_LIMITER.lock() {
        limiter.check_rate_limit("save_file_binary")?;
    } else {
        return Err("Ошибка доступа к rate limiter".into());
    }

    // Проверка размера контента
    if content.len() > MAX_FILE_SIZE {
        return Err(format!("Размер файла превышает максимальный ({} МБ)", MAX_FILE_SIZE / 1024 / 1024));
    }

    let path_buf = PathBuf::from(&path);

    // Проверка расширения файла (только .xlsx)
    if let Some(ext) = path_buf.extension() {
        let ext_str = ext.to_string_lossy().to_lowercase();
        if ext_str != "xlsx" {
            return Err("Разрешена запись только .xlsx файлов через эту команду".into());
        }
    } else {
        return Err("Файл должен иметь расширение".into());
    }

    if !is_path_allowed(&path_buf) {
        return Err("Сохранение разрешено только в папки: Загрузки, Документы или Рабочий стол".into());
    }

    std::fs::write(&path_buf, &content)
        .map_err(|e| format!("Ошибка записи: {}", e))?;

    Ok(path)
}

/// Безопасное чтение файла с проверкой пути, размера и rate limiting
#[tauri::command]
fn read_file_secure(path: String) -> Result<String, String> {
    // Rate limiting
    if let Ok(mut limiter) = RATE_LIMITER.lock() {
        limiter.check_rate_limit("read_file_secure")?;
    } else {
        return Err("Ошибка доступа к rate limiter".into());
    }
    
    let path_buf = PathBuf::from(&path);
    
    // Проверка расширения файла
    if let Some(ext) = path_buf.extension() {
        let ext_str = ext.to_string_lossy().to_lowercase();
        if ext_str != "json" && ext_str != "xml" {
            return Err("Разрешено чтение только .json и .xml файлов".into());
        }
    } else {
        return Err("Файл должен иметь расширение".into());
    }
    
    if !is_path_allowed(&path_buf) {
        return Err("Чтение разрешено только из папок: Загрузки, Документы или Рабочий стол".into());
    }
    
    // Проверяем размер файла перед чтением
    let metadata = std::fs::metadata(&path_buf)
        .map_err(|e| format!("Ошибка получения информации о файле: {}", e))?;
    
    if metadata.len() > MAX_FILE_SIZE as u64 {
        return Err(format!("Размер файла превышает максимальный ({} МБ)", MAX_FILE_SIZE / 1024 / 1024));
    }
    
    std::fs::read_to_string(&path_buf)
        .map_err(|e| format!("Ошибка чтения: {}", e))
}

/// Возвращает список разрешённых директорий
#[tauri::command]
fn get_allowed_dirs() -> Vec<String> {
    [
        dirs::download_dir(),
        dirs::document_dir(), 
        dirs::desktop_dir(),
    ]
    .into_iter()
    .flatten()
    .map(|p| p.to_string_lossy().to_string())
    .collect()
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            save_file_secure,
            save_file_binary,
            read_file_secure,
            get_allowed_dirs
        ])
        .setup(|_app| {
            // DevTools только в debug режиме
            #[cfg(debug_assertions)]
            {
                
            }
            Ok(())
        })
        .run(tauri::generate_context!())
        .expect("ошибка при запуске приложения Tauri");
}
