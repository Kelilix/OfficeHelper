use serde::{Deserialize, Serialize};
use std::process::{Child, Command};
use std::sync::Mutex;
use std::thread;
use std::time::Duration;
use std::net::TcpStream;
use tauri::{Manager, State};

// ── 后端进程状态管理 ──────────────────────────────────────────────────────

struct BackendProcess {
    child: Mutex<Option<Child>>,
}

impl BackendProcess {
    fn new() -> Self {
        Self {
            child: Mutex::new(None),
        }
    }

    fn start(&self) -> Result<(), String> {
        let exe_dir = std::env::current_exe()
            .map(|p| p.parent().unwrap().to_path_buf())
            .unwrap_or_default();

        let backend_path = Self::find_backend_path(&exe_dir);

        if !backend_path.exists() {
            return Err(format!(
                "未找到后端程序: {:?}\n\
                 搜索路径包括:\n\
                 1. dist/OfficeHelperBackend/OfficeHelperBackend.exe (开发路径)\n\
                 2. <exe>/resources/OfficeHelperBackend/OfficeHelperBackend.exe (Tauri bundle 路径)\n\
                 3. <exe>/OfficeHelperBackend/OfficeHelperBackend.exe (扁平布局)\n\
                 请确保 PyInstaller 打包已完成。\n\
                 运行: pyinstaller OfficeHelperBackend.spec --clean",
                backend_path
            ));
        }

        log::info!("启动后端: {:?}", backend_path);

        let child = Command::new(&backend_path)
            .current_dir(backend_path.parent().unwrap_or(&exe_dir))
            .spawn()
            .map_err(|e| format!("启动后端失败: {}", e))?;

        let mut guard = self.child.lock().unwrap();
        *guard = Some(child);
        Ok(())
    }

    fn stop(&self) -> Result<(), String> {
        let mut guard = self.child.lock().unwrap();
        if let Some(mut child) = guard.take() {
            let _ = child.kill();
            let _ = child.wait();
            log::info!("后端进程已停止");
        }
        Ok(())
    }

    fn is_running(&self) -> bool {
        let mut guard = self.child.lock().unwrap();
        if let Some(ref mut child) = *guard {
            child.try_wait().ok().flatten().is_none()
        } else {
            false
        }
    }

    fn is_port_open(&self) -> bool {
        TcpStream::connect_timeout(
            &"127.0.0.1:8765".parse().unwrap(),
            Duration::from_secs(1),
        ).is_ok()
    }

    fn find_backend_path(exe_dir: &std::path::Path) -> std::path::PathBuf {
        // 1. Try CARGO_MANIFEST_DIR dev path first (only set during `cargo run`)
        if let Ok(manifest_dir) = std::env::var("CARGO_MANIFEST_DIR") {
            let project_root = std::path::Path::new(&manifest_dir)
                .join("..").join("..").join("..");
            let dev_path = project_root.join("dist").join("OfficeHelperBackend")
                .join("OfficeHelperBackend.exe");
            if dev_path.exists() {
                return dev_path;
            }
        }

        // 2. Try exe_dir/resources/OfficeHelperBackend/ (Tauri bundle layout)
        let bundle_path = exe_dir.join("resources").join("OfficeHelperBackend")
            .join("OfficeHelperBackend.exe");
        if bundle_path.exists() {
            return bundle_path;
        }

        // 3. Try exe_dir/OfficeHelperBackend/ (flat release layout)
        let flat_path = exe_dir.join("OfficeHelperBackend").join("OfficeHelperBackend.exe");
        if flat_path.exists() {
            return flat_path;
        }

        // 4. Fall back to project-root dev path (common in development workflows)
        //    exe is at: front/src-tauri/target/release/office-helper-front.exe
        //    backend at: dist/OfficeHelperBackend/OfficeHelperBackend.exe
        let dev_fallback = exe_dir
            .join("..").join("..").join("..").join("..")
            .join("dist").join("OfficeHelperBackend").join("OfficeHelperBackend.exe");
        if dev_fallback.exists() {
            return dev_fallback;
        }

        // Return the expected bundle path so the error message is informative
        bundle_path
    }
}

// ── Tauri 命令 ────────────────────────────────────────────────────────────

#[derive(Debug, Serialize, Deserialize)]
struct HealthResponse {
    status: String,
    word: bool,
}

#[tauri::command]
fn start_backend(state: State<BackendProcess>) -> Result<String, String> {
    if state.is_running() {
        return Ok("already_running".to_string());
    }
    state.start()?;
    thread::sleep(Duration::from_secs(2));
    Ok("started".to_string())
}

#[tauri::command]
fn stop_backend(state: State<BackendProcess>) -> Result<(), String> {
    state.stop()
}

#[tauri::command]
fn restart_backend(state: State<BackendProcess>) -> Result<String, String> {
    state.stop()?;
    thread::sleep(Duration::from_secs(1));
    state.start()?;
    thread::sleep(Duration::from_secs(2));
    Ok("restarted".to_string())
}

#[tauri::command]
fn backend_health(state: State<BackendProcess>) -> Result<HealthResponse, String> {
    if state.is_port_open() {
        Ok(HealthResponse {
            status: "ok".to_string(),
            word: true,
        })
    } else {
        Err("后端未就绪".to_string())
    }
}

#[tauri::command]
fn is_backend_running(state: State<BackendProcess>) -> bool {
    state.is_running()
}

// ── 应用入口 ──────────────────────────────────────────────────────────────

pub fn run() {
    env_logger::Builder::from_env(
        env_logger::Env::default().default_filter_or("info")
    ).init();
    log::info!("OfficeHelper 启动中...");

    tauri::Builder::default()
        .plugin(tauri_plugin_shell::init())
        .manage(BackendProcess::new())
        .invoke_handler(tauri::generate_handler![
            start_backend,
            stop_backend,
            restart_backend,
            backend_health,
            is_backend_running,
        ])
        .setup(|app| {
            let state: State<BackendProcess> = app.state();
            if let Err(e) = state.start() {
                log::warn!("后端启动失败: {}", e);
            } else {
                thread::sleep(Duration::from_secs(2));
                log::info!("后端启动成功");
            }
            Ok(())
        })
        .on_window_event(|window, event| {
            if let tauri::WindowEvent::CloseRequested { .. } = event {
                if let Some(state) = window.try_state::<BackendProcess>() {
                    let _ = state.stop();
                }
                log::info!("应用退出");
            }
        })
        .run(tauri::generate_context!())
        .expect("运行 Tauri 应用时发生错误");
}

fn main() {
    run();
}
