pyinstaller --additional-hooks-dir=. convert.py
move dist/convert convert
rmdir /s /q build
rmdir /s /q dist
del convert.spec
