"""Run with the Python environment that has requirements.txt installed."""
from pathlib import Path
import subprocess
import sys
import shutil
import pypandoc

here = Path(__file__).resolve().parent
pandoc = Path(pypandoc.get_pandoc_path())
if not pandoc.exists(): pandoc = pandoc.with_suffix('.exe')
binary = here / 'build' / 'vendor' / 'pandoc.exe'
binary.parent.mkdir(parents=True, exist_ok=True)
shutil.copy2(pandoc, binary)
subprocess.run([sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean',
    '--windowed', '--onedir', '--name', 'MarkdownToWord',
    '--add-binary', str(binary) + ';.',
    '--distpath', str(here.parent / 'dist'),
    '--workpath', str(here / 'build'), '--specpath', str(here),
    str(here / 'app.py')], check=True)
