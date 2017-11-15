# autobugfree
transfer normal test suite (.xlsx) to a new .xlsx, which can be saved as .xml and the .xml can be imported into bugfree directly

how to use
1. add test suite file (.xlsx) into folder: src/main/resources 
2. compile source code: maven compile
3. run the code: maven exec:java -Dexec.mainClass=GenerateImportXML
4. find the result.xlsx file in folder: src/main/resources and re-save it as .xml
5. import the .xml into bugfree
cases are imported into bugfree and it is done!
