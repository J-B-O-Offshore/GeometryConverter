REM Start JBOOST
RMDIR /S /Q "Result_JBOOST_Graph"
RMDIR /S /Q "Results_JBOOST_Text"
DEL /S /Q "Log.txt"
JBOOST_v2.1.exe proj.lua
pause

