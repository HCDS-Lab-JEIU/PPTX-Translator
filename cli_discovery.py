"""Bounded local CLI discovery, including desktop launches without terminal PATH."""
import os
from pathlib import Path
import shutil

WRAPPERS={'.cmd','.bat','.ps1'}


def native(path):
    path=Path(path).expanduser()
    return path.is_file() and path.suffix.lower() not in WRAPPERS


def codex_candidates():
    found=shutil.which('codex')
    if found: yield Path(found)
    local=Path(os.environ.get('LOCALAPPDATA',Path.home()/'AppData/Local'))
    yield local/'Programs/OpenAI/Codex/bin/codex.exe'
    yield Path.home()/'.local/bin/codex'
    # VS Code and Insiders bundle a native CLI even when app PATH lacks codex.
    extensions=[]
    for directory in ('.vscode/extensions','.vscode-insiders/extensions'):
        root=Path.home()/directory
        if root.is_dir(): extensions.extend(root.glob('openai.chatgpt-*/bin/windows-*/codex.exe'))
    for path in sorted(extensions,key=lambda p:p.stat().st_mtime,reverse=True): yield path
    npm=Path(os.environ.get('APPDATA',Path.home()/'AppData/Roaming'))/'npm/node_modules'
    for package in ('@openai/codex','@openai/codex-win32-x64','@openai/codex-win32-arm64'):
        root=npm/package/'vendor'
        if root.is_dir(): yield from root.glob('*/codex/codex.exe')


def find_node():
    found=shutil.which('node')
    candidates=[Path(found)] if found else []
    for env in ('ProgramFiles','ProgramFiles(x86)','LOCALAPPDATA'):
        root=os.environ.get(env)
        if root: candidates.extend([Path(root)/'nodejs/node.exe',Path(root)/'Programs/nodejs/node.exe'])
    return next((str(p.resolve()) for p in candidates if native(p)),None)


def gemini_candidates(executable=None):
    configured=executable or os.environ.get('GEMINI_CLI_EXECUTABLE')
    if configured:
        yield Path(shutil.which(str(configured)) or configured).expanduser()
        return
    found=shutil.which('gemini')
    if found: yield Path(found)
    npm=Path(os.environ.get('APPDATA',Path.home()/'AppData/Roaming'))/'npm'
    for name in ('gemini.exe','gemini.cmd','gemini.ps1'): yield npm/name
    yield Path.home()/'.local/bin/gemini'
    for env in ('ProgramFiles','LOCALAPPDATA'):
        root=os.environ.get(env)
        if root:
            for sub in ('nodejs','Programs/nodejs'):
                yield Path(root)/sub/'node_modules/@google/gemini-cli/dist/index.js'
