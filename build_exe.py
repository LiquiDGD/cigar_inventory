import os
import shutil
import subprocess
import time
import psutil
import json
import sys

def check_requirements():
    """Check if all required packages are installed."""
    try:
        import pyinstaller
    except ImportError:
        print("PyInstaller is not installed. Installing now...")
        try:
            subprocess.run([sys.executable, "-m", "pip", "install", "pyinstaller"], check=True)
            print("PyInstaller installed successfully!")
        except Exception as e:
            print(f"Failed to install PyInstaller: {e}")
            return False
    return True

def kill_running_app():
    """Kill any running instances of the application."""
    for proc in psutil.process_iter(['pid', 'name']):
        try:
            if proc.info['name'] == 'Cigar Inventory.exe':
                proc.kill()
                time.sleep(1)  # Give the process time to terminate
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

def clean_build_files():
    """Clean up old build files."""
    paths_to_clean = ['build', 'dist', '__pycache__', 'Cigar Inventory.spec']
    
    for path in paths_to_clean:
        try:
            if os.path.isfile(path):
                os.remove(path)
            elif os.path.isdir(path):
                shutil.rmtree(path)
        except Exception as e:
            print(f"Warning: Could not remove {path}: {e}")
            time.sleep(2)  # Give Windows time to release file handles

def create_json_files():
    """Create empty JSON files if they don't exist."""
    json_files = {
        'cigar_inventory.json': [],
        'cigar_brands.json': [],
        'cigar_sizes.json': [],
        'cigar_types.json': []
    }
    
    for filename, default_data in json_files.items():
        if not os.path.exists(filename):
            with open(filename, 'w') as f:
                json.dump(default_data, f, indent=2)

def build_executable():
    """Build the executable using PyInstaller."""
    try:
        # Verify main.py exists
        if not os.path.exists('main.py'):
            print("Error: main.py not found in current directory!")
            print(f"Current directory: {os.getcwd()}")
            print("Please make sure you're running this script from the correct directory.")
            return False

        # Build command with explicit python path
        cmd = [
            sys.executable,
            '-m',
            'PyInstaller',
            '--noconfirm',
            '--onefile',
            '--windowed',
            '--icon=cigar.ico',
            '--name=Cigar Inventory',
            '--add-data=cigar_inventory.json;.',
            '--add-data=cigar_brands.json;.',
            '--add-data=cigar_sizes.json;.',
            '--add-data=cigar_types.json;.',
            'main.py'
        ]
        
        # Execute build
        result = subprocess.run(cmd, capture_output=True, text=True)
        
        # Print output for debugging
        if result.stdout:
            print("Build output:")
            print(result.stdout)
        if result.stderr:
            print("Build errors:")
            print(result.stderr)
            
        return result.returncode == 0
    except Exception as e:
        print(f"An error occurred during build: {e}")
        return False

def main():
    print("Starting build process...")
    print(f"Current working directory: {os.getcwd()}")
    
    # Check requirements
    if not check_requirements():
        print("Failed to ensure requirements. Exiting.")
        return
    
    # Kill any running instances
    print("Checking for running instances...")
    kill_running_app()
    
    # Clean up old files
    print("Cleaning up old build files...")
    clean_build_files()
    
    # Create JSON files
    print("Creating JSON files...")
    create_json_files()
    
    # Build the executable
    print("Building executable...")
    if build_executable():
        print("\nBuild completed successfully!")
        print("You can find the executable in the 'dist' folder.")
    else:
        print("\nBuild failed. Please check the error messages above.")

if __name__ == '__main__':
    main() 