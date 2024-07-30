call yarn install
@if %ERRORLEVEL% neq 0 @exit /B %ERRORLEVEL%
call yarn run dist
@if %ERRORLEVEL% neq 0 @exit /B %ERRORLEVEL%
