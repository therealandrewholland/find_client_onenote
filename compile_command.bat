python -m PyInstaller -F -i "icon.ico" --name="Find Client OneNote" "%~dp0Find Client OneNote.pyw" --onedir

copy "icon.ico" "dist\Find Client OneNote\" & copy "transparent.ico" "dist\Find Client OneNote\" & copy "IT Glue Client IDs_JAN2024.json" "dist\Find Client OneNote\"
