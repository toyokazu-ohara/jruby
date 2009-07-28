@echo off

rem This script just makes jruby.bat run with javaw

setlocal
set JAVA_COMMAND=javaw
call "%~dp0_jrubyvars" %*

if %JRUBY_BAT_ERROR%==0 "%_STARTJAVA%" %_VM_OPTS% -Xbootclasspath/a:"%JRUBY_CP%" -classpath "%CP%;%CLASSPATH%" -Djruby.home="%JRUBY_HOME%" -Djruby.lib="%JRUBY_HOME%\lib" -Djruby.shell="cmd.exe" -Djruby.script=jruby.bat -Djacob.dll.name.x86="%JRUBY_HOME%\lib\native\x86-win\jacob-1.14.3-x86.dll" -Djacob.dll.name.x64="%JRUBY_HOME%\lib\native\x64-win\jacob-1.14.3-x64.dll" org.jruby.Main %JRUBY_OPTS% %_RUBY_OPTS%
set E=%ERRORLEVEL%

call "%~dp0_jrubycleanup"
endlocal