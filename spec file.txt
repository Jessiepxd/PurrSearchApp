#!/usr/bin/env python       # Shebang for VSCode
# -*- mode: python -*-      # File variable for EMACS.
# vim: set syntax=python:   # Modeline for VIM
# Note: If build doesn't work, try the following command:
#   pip install --force-reinstall --no-binary :all: pyinstaller
# https://github.com/pyinstaller/pyinstaller/issues/3806
import os, sys, string
from PyInstaller.utils.hooks import collect_submodules
from PyInstaller.utils.hooks import collect_data_files

# Extend PYTHONPATH with the script path.
SCRIPTPATH = DISTPATH.partition('dist')[0]
sys.path.append(SCRIPTPATH)

import PurrSearch

# Generate versioninfo.txt
template = open(os.path.join(SCRIPTPATH, 'version-template.txt'), 'r').read()
info = string.Template(template)
infofile = open("versioninfo.txt", 'w')
infofile.write(info.substitute(PurrSearch.VERSIONINFO))
infofile.close()


a = Analysis(
    scripts=['PurrSearch.py'],
    pathex=[],                      # Paths to be searched before sys.path.
    binaries=[],                    # Additional binaries, DLLs, etc.
    datas=[('images/splash.png', 'images')],                         # Additional data files to include.
    hiddenimports=[
        'wx._xml',
        'wx.xml',
        'wx.richtext', 
        'wx.html',
        'wx.adv',
        'wx.lib.throbber',
        'wx.lib.mixins.inspection'
    ],               # Additional modules to include.
    hookspath=[],                   # Additional module hooks.
    hooksconfig={},
    runtime_hooks=[],               # Script filenames to run at startup.
    excludes=[],
    noarchive=False,                # If True keep files individual.
)

# Create a ZlibArchive that contains all pure Python modules.
pyz = PYZ(a.pure, a.zipped_data)

# Create the final executable of the frozen app.
exe = EXE(
    pyz,                            # Target - Pure-python libraries.
    a.scripts,                      # TOC - Main script(s) to execute.
    a.binaries,  
    a.zipfiles, 
    a.datas,                        # TOC - Data files.
    [],
    name='PurrSearch',              # The filename for the executable.
    debug=False,                    # Show progress messages or not.
    uac_admin=False,
    version="versioninfo.txt",      # Add version resource (Windows only).
    bootloader_ignore_signals=False,
    strip=False,
    runtime_tmpdir=None,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['PurrIcon.ico'],              # Add icon resource (Windows/MAC only).
    console=False,                  # Run in Windowed mode (Windows/MAC only).
    splash='images/splash.png',  # Add the splash screen here
)





