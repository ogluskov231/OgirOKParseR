# Підготовка Windows-релізу

## 1. Збірка програми
На Windows у корені проєкту запустити:

`scripts\build_exe_windows.bat`

Результат:

`dist\OgrOK_UA\OgrOK_UA.exe`

## 2. Створення установника
Потрібен Inno Setup.

Відкрити:

`installer\OgrOK_UA.iss`

і виконати Compile.

Готовий установник з'явиться в:

`installer_output\OgrOK_UA_Setup_v0.6.exe`

## Що навмисно не входить до чистого пакета
- `.venv`
- `build`
- `dist`
- `.pytest_cache`
- `__pycache__`
- старі результати `output`
- документи користувача из `input\documents`
- старий Word-еталон, який був потрібен лише для генерації mapping.json
- Linux/macOS та старий Windows launcher
