@echo off
rem Pick one of the following for testing 
rem IRAhandle_tweets_1_15.csv
rem IRAhandle_tweets_1_100.csv
rem IRAhandle_tweets_1_1000.csv
rem IRAhandle_tweets_1_243892.csv

set dfile="excel-integrator-test-data.csv"
set sfile=

echo  Pick one of the following data files
echo .
echo    1 IRAhandle_tweets_1_15.csv
echo    2 IRAhandle_tweets_1_100.csv
echo    3 IRAhandle_tweets_1_1000.csv
echo    4 IRAhandle_tweets_1_243892.csv
echo .

set /p input="Enter ordinal of dataset to use; 1, 2, 3, or 4: "


if %input%==1 set sfile="IRAhandle_tweets_1_15.csv"
if %input%==2 set sfile="IRAhandle_tweets_1_100.csv"
if %input%==3 set sfile="IRAhandle_tweets_1_1000.csv"
if %input%==4 set sfile="IRAhandle_tweets_1_243892.csv"
if [%sfile%]==[] GOTO ERROR

echo .. Copy %sfile% to %dfile% 
copy %sfile% %dfile% 
timeout 5
exit 

ERROR: 
echo "** Select a valid option" 
timeout 10 
pause 
exit 
