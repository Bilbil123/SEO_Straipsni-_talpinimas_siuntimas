import os
import subprocess

def create_icon():
    # Create iconset directory
    os.makedirs('app_icon.iconset', exist_ok=True)
    
    # Convert PNG to different sizes
    sizes = [16, 32, 64, 128, 256, 512, 1024]
    for size in sizes:
        subprocess.run([
            'sips',
            '-z', str(size), str(size),
            'icon.png',
            '--out', f'app_icon.iconset/icon_{size}x{size}.png'
        ])
    
    # Create icns file
    subprocess.run([
        'iconutil',
        '-c', 'icns',
        'app_icon.iconset'
    ])
    
    # Clean up
    subprocess.run(['rm', '-rf', 'app_icon.iconset'])
    
    print("Icon created successfully!")

if __name__ == "__main__":
    create_icon() 