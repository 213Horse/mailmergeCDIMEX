#!/usr/bin/env python3
"""
Auto-installer for Bookmedi Mail Merge
Chạy script này để tự động cài đặt môi trường và dependencies
"""

import os
import sys
import subprocess
import platform
from pathlib import Path

def run_command(cmd, check=True):
    """Run a command and return the result"""
    print(f"Running: {cmd}")
    try:
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if check and result.returncode != 0:
            print(f"Error: {result.stderr}")
            return False
        return True
    except Exception as e:
        print(f"Error running command: {e}")
        return False

def check_python():
    """Check if Python 3.8+ is available"""
    print("Checking Python version...")
    try:
        version = sys.version_info
        if version.major >= 3 and version.minor >= 8:
            print(f"✓ Python {version.major}.{version.minor}.{version.micro} found")
            return True
        else:
            print(f"✗ Python version {version.major}.{version.minor} is too old. Need 3.8+")
            return False
    except Exception as e:
        print(f"✗ Error checking Python: {e}")
        return False

def install_python():
    """Install Python if not available"""
    system = platform.system().lower()
    
    if system == "darwin":  # macOS
        print("Installing Python on macOS...")
        if run_command("which brew"):
            return run_command("brew install python@3.13")
        else:
            print("Please install Homebrew first: https://brew.sh")
            return False
    elif system == "windows":
        print("Please download and install Python from: https://python.org")
        return False
    else:  # Linux
        print("Installing Python on Linux...")
        return run_command("sudo apt-get update && sudo apt-get install -y python3 python3-pip python3-venv")

def create_venv():
    """Create virtual environment"""
    print("Creating virtual environment...")
    if Path(".venv").exists():
        print("✓ Virtual environment already exists")
        return True
    
    return run_command(f"{sys.executable} -m venv .venv")

def install_dependencies():
    """Install required packages"""
    print("Installing dependencies...")
    
    # Determine the correct pip path
    if platform.system().lower() == "windows":
        pip_path = ".venv\\Scripts\\pip"
    else:
        pip_path = ".venv/bin/pip"
    
    # Upgrade pip first
    run_command(f"{pip_path} install --upgrade pip")
    
    # Install required packages
    packages = [
        "pandas",
        "openpyxl", 
        "streamlit",
        "requests"
    ]
    
    for package in packages:
        print(f"Installing {package}...")
        if not run_command(f"{pip_path} install {package}"):
            print(f"Failed to install {package}")
            return False
    
    print("✓ All dependencies installed successfully")
    return True

def create_launcher_scripts():
    """Create easy-to-use launcher scripts"""
    print("Creating launcher scripts...")
    
    # Windows batch file
    if platform.system().lower() == "windows":
        with open("run_gui.bat", "w") as f:
            f.write("""@echo off
echo Starting Bookmedi Mail Merge...
call .venv\\Scripts\\activate
streamlit run streamlit_app.py
pause
""")
        print("✓ Created run_gui.bat")
    
    # Unix shell script
    else:
        with open("run_gui.sh", "w") as f:
            f.write("""#!/bin/bash
echo "Starting Bookmedi Mail Merge..."
source .venv/bin/activate
streamlit run streamlit_app.py
""")
        # Make executable
        os.chmod("run_gui.sh", 0o755)
        print("✓ Created run_gui.sh")

def main():
    print("=" * 50)
    print("Bookmedi Mail Merge - Auto Installer")
    print("=" * 50)
    
    # Check Python
    if not check_python():
        print("\nPython not found or version too old.")
        response = input("Would you like to install Python? (y/n): ").lower()
        if response == 'y':
            if not install_python():
                print("Failed to install Python. Please install manually.")
                return
        else:
            print("Please install Python 3.8+ manually and run this script again.")
            return
    
    # Create virtual environment
    if not create_venv():
        print("Failed to create virtual environment.")
        return
    
    # Install dependencies
    if not install_dependencies():
        print("Failed to install dependencies.")
        return
    
    # Create launcher scripts
    create_launcher_scripts()
    
    print("\n" + "=" * 50)
    print("Installation completed successfully!")
    print("=" * 50)
    print("\nTo run the application:")
    
    if platform.system().lower() == "windows":
        print("1. Double-click 'run_gui.bat'")
    else:
        print("1. Double-click 'run_gui.sh' or run: ./run_gui.sh")
    
    print("2. Wait for the browser to open automatically")
    print("3. If browser doesn't open, go to: http://localhost:8501")
    print("\nNote: Keep the terminal window open while using the application.")

if __name__ == "__main__":
    main()
