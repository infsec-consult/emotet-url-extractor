Extracts the hardcoded URLs from an emotet .doc 

Usage:  
python3 ./emotet_extractor.py $file

-h for help  
-s will dump the script after base64-decode
-g If the normal search cant find a script, try -g. Warning, long runtime


And for adjusting to new emotet versions:  
-d lets you chose a delimiter for splitting the obfuscated script

  
NEW: Now needs oledump.py in the same folder
https://blog.didierstevens.com/programs/oledump-py/
