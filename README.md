# ExcelUtils
A simple java static utility class for reading and writing excel files and validating header rows.
## Getting Started
Simply include the the java file in your project, add package declaration.

### For reading excel files
Simply provide multipart file and class instance as parameter to excelConverter method, It will return List of objects of class instance.
```
List<yourClassInstance> = ExcelUtils.excelConverter(yourFile, yourClassInstance);
```
The excel sheet's column header should match with class's attributes. Headernames could be space seperated, hyphen, forward slash or underscore.
Class should have corresponding camel case variables of column header names.
```
//If excel is having one column with header name as "First Name", "First-Name", "First/Name", "First_Name" then yourClassInstance class should have attribute as follows.
private String firstName;
```

### For writing to excel files
Simply provide Map of sheetname and List of your class instance objects, Map of sheetname and List of header name keys and Map of header name keys and corresponding local labels.
```
File  = ExcelUtils.generateExcel(yourMapOfSheetNameAndListOfClassInstanceObjects, yourMapOfSheetNameAndListOfHeaderNames, yourMapOfHeaderKeysAndActualHeaderNames);
```
Header names should match with attributes of class.


### Example
will be added soon.
