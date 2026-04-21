@echo off
setlocal
set "ROOT=%~dp0"
pushd "%ROOT%"
python ".\email_template_gui.py"
popd
endlocal
