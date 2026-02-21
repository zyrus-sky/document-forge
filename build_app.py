import os
import subprocess
import shutil
import sys

def build_svelte():
    print("🚀 Building SvelteKit frontend...")
    frontend_dir = os.path.join(os.getcwd(), 'frontend')
    
    # Run npm run build
    try:
        if sys.platform == "win32":
            subprocess.run(["npm.cmd", "run", "build"], cwd=frontend_dir, check=True)
        else:
            subprocess.run(["npm", "run", "build"], cwd=frontend_dir, check=True)
        print("✅ Frontend build complete.\n")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build frontend: {e}")
        sys.exit(1)

def build_pyinstaller():
    print("📦 Packaging with PyInstaller...")
    
    # Path to the Svelte build folder
    frontend_build = os.path.join('frontend', 'build')
    
    # Construct PyInstaller command
    cmd = [
        "pyinstaller",
        "--noconfirm",
        "--windowed", # Don't show terminal window
        "--name", "DocumentForge",
        "--add-data", f"{frontend_build};frontend/build", # Include svelte files
        "app.py"
    ]
    
    try:
        subprocess.run(cmd, check=True)
        print("✅ PyInstaller build complete. Executable is in the 'dist' folder!\n")
    except subprocess.CalledProcessError as e:
        print(f"❌ Failed to build executable: {e}")
        sys.exit(1)

if __name__ == "__main__":
    print(f"--- Document Forge Build Script ({sys.platform}) ---")
    
    # Ensure dependencies are locally installed
    build_svelte()
    build_pyinstaller()
    
    print("🎉 All done! Check the 'dist/DocumentForge' folder for your executable.")
