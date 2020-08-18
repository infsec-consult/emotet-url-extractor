Extracts the hardcoded URLs from an emotet .doc 

Usage:
python3 ./emotet_extractor.py $file

-h for help
-s will dump the script after base64-decode


And for adjusting to new emotet versions:
-i prints the obfuscated script
-d lets you chose a delimiter for splitting the obfuscated script
