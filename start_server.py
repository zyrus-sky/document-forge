import subprocess
import webbrowser
import time
import os
import sys

def main():
    print("🚀 Starting Document Forge Web App...")
    
    # Start Backend
    backend_process = subprocess.Popen(
        [sys.executable, "backend.py"],
        cwd=os.getcwd()
    )
    
    # Start Frontend
    frontend_dir = os.path.join(os.getcwd(), "frontend")
    npm_cmd = "npm.cmd" if sys.platform == "win32" else "npm"
    frontend_process = subprocess.Popen(
        [npm_cmd, "run", "dev", "--", "--port", "5173", "--open"],
        cwd=frontend_dir
    )
    
    print("✅ Servers are booting up!")
    print("Backend: http://localhost:8000")
    print("Frontend: http://localhost:5173")
    print("\nPress Ctrl+C to stop both servers.")
    
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nStopping servers...")
        backend_process.terminate()
        frontend_process.terminate()
        sys.exit(0)

if __name__ == "__main__":
    main()
