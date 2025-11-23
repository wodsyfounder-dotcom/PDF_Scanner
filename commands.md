```powershell
$env:DEBUG_MODE="1"; python -m ui_next.qt_main 2>&1 | Tee-Object -FilePath debug_output_new.txt
```
