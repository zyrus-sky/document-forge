"""
Document Forge — Build Script
Builds the SvelteKit frontend, packages with PyInstaller, 
and optionally compiles the Inno Setup installer.
"""
import os
import subprocess
import shutil
import sys

APP_NAME = "DocumentForge"
APP_VERSION = "1.0.0"
FRONTEND_DIR = os.path.join(os.getcwd(), "frontend")
DIST_DIR = os.path.join(os.getcwd(), "dist", APP_NAME)

def build_svelte():
    print("═" * 60)
    print("  STEP 1: Building SvelteKit Frontend")
    print("═" * 60)
    npm_cmd = "npm.cmd" if sys.platform == "win32" else "npm"
    try:
        subprocess.run([npm_cmd, "run", "build"], cwd=FRONTEND_DIR, check=True)
        print("✅ Frontend build complete.\n")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build frontend: {e}")
        sys.exit(1)

def build_pyinstaller():
    print("═" * 60)
    print("  STEP 2: Packaging with PyInstaller")
    print("═" * 60)
    
    # Kill any running instances first
    print("  Closing any running DocumentForge instances...")
    subprocess.run(["taskkill", "/F", "/IM", f"{APP_NAME}.exe"], 
                   capture_output=True)
    
    # Clean old dist to avoid permission errors
    if os.path.exists(DIST_DIR):
        import time
        time.sleep(1)  # Give Windows time to release file handles
        try:
            shutil.rmtree(DIST_DIR)
            print("  Cleaned old build artifacts.")
        except PermissionError:
            print("⚠️  Could not remove old dist folder. Close DocumentForge.exe and try again.")
            sys.exit(1)

    frontend_build = os.path.join("frontend", "build")

    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--windowed",
        "--name", APP_NAME,
        "--icon", "installer_assets/icon.ico" if os.path.exists("installer_assets/icon.ico") else "NONE",
        # Bundle frontend static files
        "--add-data", f"{frontend_build};frontend/build",
        # Bundle backend.py as a data file so it can be imported
        "--add-data", "backend.py;.",
        # Hidden imports for FastAPI + dependencies
        "--hidden-import", "uvicorn",
        "--hidden-import", "uvicorn.logging",
        "--hidden-import", "uvicorn.loops",
        "--hidden-import", "uvicorn.loops.auto",
        "--hidden-import", "uvicorn.protocols",
        "--hidden-import", "uvicorn.protocols.http",
        "--hidden-import", "uvicorn.protocols.http.auto",
        "--hidden-import", "uvicorn.protocols.websockets",
        "--hidden-import", "uvicorn.protocols.websockets.auto",
        "--hidden-import", "uvicorn.lifespan",
        "--hidden-import", "uvicorn.lifespan.on",
        "--hidden-import", "fastapi",
        "--hidden-import", "fastapi.middleware.cors",
        "--hidden-import", "starlette",
        "--hidden-import", "starlette.staticfiles",
        "--hidden-import", "starlette.responses",
        "--hidden-import", "starlette.routing",
        "--hidden-import", "starlette.middleware",
        "--hidden-import", "starlette.middleware.cors",
        "--hidden-import", "anyio",
        "--hidden-import", "anyio._backends",
        "--hidden-import", "anyio._backends._asyncio",
        "--hidden-import", "pydantic",
        "--hidden-import", "multipart",
        "--hidden-import", "python-multipart",
        "--hidden-import", "pdfplumber",
        "--hidden-import", "tabula",
        "--hidden-import", "openpyxl",
        "--hidden-import", "pandas",
        "--hidden-import", "docx",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--hidden-import", "webview",
        "--hidden-import", "clr_loader",
        "--hidden-import", "webview.platforms.winforms",
        "--hidden-import", "mimetypes",
        "--hidden-import", "email.mime",
        # Collect all needed packages
        "--collect-all", "pdfplumber",
        "--collect-all", "tabula",
        "--collect-all", "webview",
        "--collect-all", "uvicorn",
        "app.py"
    ]

    # Remove invalid --icon if no icon file exists
    if not os.path.exists("installer_assets/icon.ico"):
        cmd = [c for c in cmd if c != "NONE" and c != "--icon"]

    try:
        subprocess.run(cmd, check=True)
        print(f"✅ PyInstaller build complete → dist/{APP_NAME}/\n")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build executable: {e}")
        sys.exit(1)

def compile_installer():
    print("═" * 60)
    print("  STEP 3: Compiling Inno Setup Installer")
    print("═" * 60)

    iss_file = "installer.iss"
    if not os.path.exists(iss_file):
        print("⚠️  installer.iss not found — skipping installer compilation.")
        return

    # Try common Inno Setup paths
    iscc_paths = [
        r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe",
        r"C:\Program Files\Inno Setup 6\ISCC.exe",
        r"C:\Program Files (x86)\Inno Setup 5\ISCC.exe",
    ]

    iscc = None
    for path in iscc_paths:
        if os.path.exists(path):
            iscc = path
            break

    if not iscc:
        # Try PATH
        result = shutil.which("ISCC")
        if result:
            iscc = result

    if iscc:
        try:
            subprocess.run([iscc, iss_file], check=True)
            print("✅ Installer compiled successfully!\n")
        except subprocess.CalledProcessError as e:
            print(f"❌ Inno Setup compilation failed: {e}")
    else:
        print("⚠️  Inno Setup (ISCC.exe) not found. Install from https://jrsoftware.org/isdl.php")
        print(f"   You can manually compile: ISCC.exe {iss_file}")

if __name__ == "__main__":
    print(f"\n{'═' * 60}")
    print(f"  Document Forge Build Pipeline v{APP_VERSION}")
    print(f"  Platform: {sys.platform}")
    print(f"{'═' * 60}\n")

    # build_svelte()  # Bypassed for quicker Pyinstaller testing
    build_pyinstaller()
    compile_installer()

    print("═" * 60)
    print("  🎉 BUILD COMPLETE!")
    print(f"  Executable: dist/{APP_NAME}/{APP_NAME}.exe")
    if os.path.exists(f"installer_output/{APP_NAME}_Setup.exe"):
        print(f"  Installer:  installer_output/{APP_NAME}_Setup.exe")
    print("═" * 60)
