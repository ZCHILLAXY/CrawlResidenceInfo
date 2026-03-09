"""
Pack - PyInstaller
"""
import os
import sys
import shutil
from pathlib import Path

def main():
    """run pack"""
    print("=" * 60)
    print("居住证查询工具 - 打包")
    print("Made with ❤️by Z🐻")
    print("=" * 60)

    # Check PyInstaller status
    try:
        import PyInstaller
    except ImportError:
        print("Error: No PyInstaller")
        print("Run: pip install pyinstaller")
        return 1

    # Clear up old build files
    print("\nClear up old build files...")
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        if Path(dir_name).exists():
            shutil.rmtree(dir_name)
            print(f"  Deleted: {dir_name}")

    # Build PyInstaller command
    print("Start packing...")

    # Use spec to pack
    if Path('build.spec').exists():
        cmd = 'pyinstaller build.spec --clean'
    else:
        # if no spec file, use default param
        cmd = (
            'pyinstaller '
            '--name "居住证查询工具" '
            '--onefille '
            '--windowed '
            '--clean '
            '--noconfirm '
            'gui.py'
        )

        print(f"Execute: {cmd}")
        result = os.system(cmd)

        if result == 0:
            print("\n" + "=" * 60)
            print("Packed Successfully!")
            print("=" * 60)
            print(f"\nExecutable file location: {Path('dist').absolute()}")
            print("\nInstruction:")
            print("1. dist可执行文件复制到目标电脑")
            print("2. 安装Tesseract-OCR")
            print("3. 双击运行程序")
            return 0
        else:
            print("\nPacked Failed, please check the error infos")
            return 1

if __name__ == '__main__':
    sys.exit(main())
