#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
构建脚本
将程序打包成独立的 Windows 可执行文件

使用方法：
    python build.py

构建完成后，可执行文件在 dist/通达信定时下载工具/ 目录下
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path


def check_python():
    """检查 Python 版本"""
    if sys.version_info < (3, 8):
        print("❌ 需要 Python 3.8 或更高版本")
        return False
    print(f"✅ Python {sys.version_info.major}.{sys.version_info.minor}")
    return True


def check_dependencies():
    """检查依赖是否安装"""
    required = ['pyautogui', 'pywin32', 'apscheduler']
    missing = []

    for module in required:
        try:
            __import__(module)
        except ImportError:
            missing.append(module)

    if missing:
        print(f"❌ 缺少依赖：{', '.join(missing)}")
        print("请运行：pip install -r requirements.txt")
        return False

    print("✅ 所有依赖已安装")
    return True


def clean():
    """清理构建目录"""
    dirs_to_clean = ['build', 'dist']

    for d in dirs_to_clean:
        if Path(d).exists():
            print(f"清理 {d}/ ...")
            shutil.rmtree(d)

    # 清理 spec 文件生成的额外文件
    spec = Path("tongdaxin_tool.spec")
    if spec.exists():
        spec.unlink()

    print("✅ 清理完成")


def build():
    """执行构建"""
    print("\n" + "=" * 50)
    print("开始构建...")
    print("=" * 50)

    # 检查 PyInstaller
    try:
        import PyInstaller
        print(f"✅ PyInstaller {PyInstaller.__version__}")
    except ImportError:
        print("❌ PyInstaller 未安装")
        print("请运行：pip install pyinstaller")
        return False

    # 执行 PyInstaller
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--name=通达信定时下载工具',
        '--onefile',
        '--windowed',
        '--clean',
        '--distpath=dist',
        '--workpath=build',
        '--specpath=.',
        '--add-data=config.ini;.',
        'main.py'
    ]

    print(f"\n执行命令：{' '.join(cmd)}")
    print("-" * 50)

    try:
        result = subprocess.run(cmd, check=True)
        print("\n✅ 构建成功！")
        print("\n输出目录：dist/通达信定时下载工具.exe")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n❌ 构建失败：{e}")
        return False


def make_installer():
    """
    制作分发包
    注意：这个函数创建一个 ZIP 压缩包
    真正的安装程序需要使用 NSIS、Inno Setup 等工具
    """
    dist_dir = Path('dist')
    output_dir = Path('release')

    if not dist_dir.exists():
        print("❌ 找不到 dist 目录，请先执行构建")
        return False

    output_dir.mkdir(exist_ok=True)

    # 创建 ZIP 包
    import zipfile

    zip_name = output_dir / '通达信定时下载工具_v1.0.zip'
    source_dir = dist_dir / '通达信定时下载工具'

    if not source_dir.exists():
        source_dir = dist_dir / '通达信定时下载工具.exe'
        if not source_dir.exists():
            print("❌ 找不到生成的可执行文件")
            return False

    print(f"\n创建分发包：{zip_name}")

    with zipfile.ZipFile(zip_name, 'w', zipfile.ZIP_DEFLATED) as zipf:
        if source_dir.is_dir():
            for file in source_dir.rglob('*'):
                if file.is_file():
                    arcname = file.relative_to(source_dir.parent)
                    zipf.write(file, arcname)
                    print(f"  添加：{arcname}")
        else:
            arcname = source_dir.name
            zipf.write(source_dir, arcname)
            print(f"  添加：{arcname}")

    print(f"\n✅ 分发包已创建：{zip_name}")
    return True


def main():
    print("=" * 50)
    print("通达信定时下载工具 - 构建脚本")
    print("=" * 50)

    if not check_python():
        sys.exit(1)

    if len(sys.argv) > 1:
        command = sys.argv[1].lower()

        if command in ['clean', 'c']:
            clean()
        elif command in ['rebuild', 'r']:
            clean()
            build()
        elif command in ['release', 'dist']:
            make_installer()
        else:
            print(f"未知命令：{command}")
            print("可用命令：clean, rebuild, release")
    else:
        # 默认：清理 + 构建
        if check_dependencies():
            clean()
            if build():
                # 询问是否创建分发包
                response = input("\n是否创建分发包？(y/n): ")
                if response.lower() == 'y':
                    make_installer()


if __name__ == '__main__':
    main()
