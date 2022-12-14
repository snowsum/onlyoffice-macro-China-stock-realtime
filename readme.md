This is a macro script to get realtime stock price from website API when you open the sheet

This script do 4 things. 
- Get stock number from the table, A1:A200
- Get stock price info from website API
- Processing strings received from web API
- Fill the processed info into table B1:B200

Build this script under the help of ChatGPT with little code experience.

Note
-the stocknumber should be in format "sh600519" "hk00700" "sz002304"
 sh/hk/sz is the short code for largest 3 stock exchange
-the script should be run in certain sheet or there will be an error, in this script it is in sheet4
-the string received in format like 
v_s_sh600519="1~贵州茅台~600519~1758.15~8.15~0.47~9114~160244~~22085.84~GP-A";
we use function to get info between the "~ ~"

