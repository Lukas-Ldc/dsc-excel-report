# Desired State Configuration Excel Report

This report/dashboard allows you to keep an eye on the status of your configurations. The script will create an XLSX file and then fill it as time goes by. The file will contain several spreadsheets. The first one, named "Global", contains a summary of the Desired States Configurations of the nodes given by "Test-DscConfiguration". The other spreadsheets are linked to each node and contains the status given by "Get-DscConfigurationStatus" and "Test- DscConfiguration" globally and for each resource. The color code is: Green (Correct), Red (Incorrect) and Blue (Corrected). The corrected state means that the element was incorrect, but thanks to Start-DscConfiguration, it has become correct. The date of the script execution is present in the first column. To run the script, the Microsoft Excel software must be installed. The nodes to be analyzed must be added in the "nodes.txt" file separated by commas. To activate the correction, the "-Correct" parameter must be specified when launching the script.

This is how the global spreadsheet looks like :

<img src="/img.readme/excel_global.png?raw=true" alt="The global spreadsheet" width="600">

This is how a node spreadsheet looks like :

<img src="/img.readme/excel_node.png?raw=true" alt="The node spreadsheet" width="600">

(FAUX/VRAI meaning FALSE/TRUE will appear in your server language)
