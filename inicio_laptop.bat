@echo off
cd /d "%~dp0"

powershell -NoProfile -ExecutionPolicy Bypass -Command "$log='update.log'; Add-Content $log ('==== INICIO ' + (Get-Date) + ' ===='); $out='git_pull.out'; $err='git_pull.err'; Remove-Item $out,$err -ErrorAction SilentlyContinue; $p=Start-Process -FilePath 'git' -ArgumentList 'pull origin main' -WorkingDirectory (Get-Location) -NoNewWindow -PassThru -RedirectStandardOutput $out -RedirectStandardError $err; if ($p.WaitForExit(12000)) { if (Test-Path $out) { Add-Content $log (Get-Content $out -Raw) }; if (Test-Path $err) { Add-Content $log (Get-Content $err -Raw) } } else { try { $p.Kill() } catch {}; Add-Content $log 'Git pull cancelado por timeout de 12 segundos.' }; Remove-Item $out,$err -ErrorAction SilentlyContinue"

start "" pythonw "registro_laptop.pyw"