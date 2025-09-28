import subprocess
import sys

def install_requirements():
    print("Installing required packages...")
    try:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', '-r', 'requirements.txt'])
        print("All packages installed successfully!")
    except subprocess.CalledProcessError:
        print("Error installing packages. Please check your internet connection or try installing them manually.")
        sys.exit(1)

if __name__ == "__main__":
    install_requirements()