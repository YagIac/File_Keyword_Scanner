# -*- mode: python ; coding: utf-8 -*-


a = Analysis(
    ['./智能文件关键词扫描工具.py'],
    pathex=[],
    binaries=[],
    datas=[('D:/PythonProject/project_2/.venv/Lib/site-packages/customtkinter', 'customtkinter'),('D:/PythonProject/project_2/.venv/Lib/site-packages/rapidocr_onnxruntime/config.yaml',  'rapidocr_onnxruntime'),('D:/PythonProject/project_2/.venv/Lib/site-packages/rapidocr_onnxruntime/models', 'rapidocr_onnxruntime/models'),('D:/tools/Release-26.02.0-0/poppler-26.02.0/Library/bin', 'poppler/bin'), ('D:/tools/Release-26.02.0-0/poppler-26.02.0/Library/lib', 'poppler/lib')],
    hiddenimports=['rapidocr_onnxruntime', 'onnxruntime', 'docx', 'openpyxl', 'pptx', 'pdfplumber', 'pdf2image'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='智能文件关键词扫描工具',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)

coll = COLLECT(
    exe,
    a.binaries,      
    a.datas,       
    strip=False,
    upx=True,
    name='智能文件关键词扫描工具',
)
