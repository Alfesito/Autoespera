@echo off
REM Intenta con python3
python3 Autoespera.py
IF %ERRORLEVEL% NEQ 0 (
    REM Si falla, intenta con python
    python Autoespera.py
    IF %ERRORLEVEL% NEQ 0 (
        REM Si vuelve a fallar, intenta con la ruta relativa
        python .\Autoespera.py
        IF %ERRORLEVEL% NEQ 0 (
            REM Si vuelve a fallar, intenta con la ruta relativa
            .\Autoespera.py
            python .\Autoespera.py
            IF %ERRORLEVEL% NEQ 0 (
                echo Error: No se pudo ejecutar Autoespera.py
                pause
                exit /b 1
            )
        )
    )
)
pause
