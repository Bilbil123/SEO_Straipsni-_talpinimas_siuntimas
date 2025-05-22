import os
import subprocess
import sys
import shutil

def build_app():
    # Install required packages if not already installed
    try:
        import PyInstaller
    except ImportError:
        print("Installing PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
    
    # Clean up previous builds
    if os.path.exists('build'):
        shutil.rmtree('build')
    if os.path.exists('dist'):
        shutil.rmtree('dist')
    
    # Create the app bundle
    print("Building Knygospigiau Newsletter.app...")
    try:
        subprocess.check_call([
            "pyinstaller",
            "--clean",
            "knygospigiau_newsletter.spec"
        ])
        
        print("\nBuild completed successfully!")
        print("The app bundle can be found in the 'dist' directory.")
        
        # Copy the app to Applications folder
        app_path = os.path.join('dist', 'Knygospigiau Newsletter.app')
        if os.path.exists(app_path):
            print("\nWould you like to copy the app to your Applications folder? (y/n)")
            if input().lower() == 'y':
                dest_path = '/Applications/Knygospigiau Newsletter.app'
                if os.path.exists(dest_path):
                    shutil.rmtree(dest_path)
                shutil.copytree(app_path, dest_path)
                print(f"App copied to {dest_path}")
    except subprocess.CalledProcessError as e:
        print(f"\nBuild failed with error: {e}")
        sys.exit(1)
    except PermissionError:
        print("\nPermission denied. Please run the script with sudo or copy the app manually from the 'dist' folder.")

if __name__ == "__main__":
    build_app() 