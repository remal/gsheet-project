call build
@if %ERRORLEVEL% neq 0 @exit /B %ERRORLEVEL%
call git add --all .
@if %ERRORLEVEL% neq 0 @exit /B %ERRORLEVEL%
call git commit --message="%TIME%"
@if %ERRORLEVEL% neq 0 @exit /B %ERRORLEVEL%
call git push
@if %ERRORLEVEL% neq 0 @exit /B %ERRORLEVEL%
