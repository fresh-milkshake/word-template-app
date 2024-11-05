@ECHO OFF

SET dist_path=bin
SET name=word-template-app
SET build_path=build

ECHO Установка зависимостей...

pip install -r requirements.txt > nul
pip install pyinstaller > nul

ECHO Создание исполняемого файла...

pyinstaller mainwindow.py --onefile --noconsole --name %name% --distpath %dist_path% --workpath %build_path% --log-level ERROR

IF %ERRORLEVEL% NEQ 0 (
    echo Произошла ошибка! Попробуйте запустить скрипт еще раз.
    PAUSE
    EXIT
)

ECHO Исполняемый файл создан.

ECHO Очистка ненужных файлов...

DEL %name%.spec
RD /s /q %build_path%

SET /P "should_run=Запустить скомпилированный файл? (y/n): "
IF "%should_run%"=="y" (
    ECHO Запуск .exe файла...
    %dist_path%\%name%
)

ECHO Скрипт завершен.
PAUSE