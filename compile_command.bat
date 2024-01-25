python -m PyInstaller -F -i "icon.ico" --name="Find Client OneNote" "%~dp0Find Client OneNote.pyw" --onedir

copy "icon.ico" "dist\Find Client OneNote\" & copy "transparent.ico" "dist\Find Client OneNote\" & copy "chromedriver.exe" "dist\Find Client OneNote\" & copy "get_itglue_ids.js" "dist\Find Client OneNote\" & copy "get_itglue_ids.py" "dist\Find Client OneNote\"