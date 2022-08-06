## **# CyberAttackSequencingTool**
### **Description**
Network attack records are often complex and redundant, and a large number of attacks from the same IP take up a lot of space in the records, which can actually be replaced by just one integrated piece of information.

CyberAttackSequencingTool is used to organize this data. You can choose a category, and the program will automatically sort the records according to that category. You can also choose to collapse the information of the same category into one, and you can choose to Repeated parts are omitted.
### **Usage：**
First, please make sure that your computer has a running environment of Java 17 or above.

You may be able to find a suitable download on [Java Downloads | Oracle](https://www.oracle.com/java/technologies/downloads/#jdk18-linux).

Next, download the latest release version.

As this is a command line interface program, please type the command exactly in the following format on the command line.

`	java -jar xxx.jar [-option1 [args]] [-option2 [args]] ……  `

Where xxx.jar is the file name, which you downloaded just now.

You may need to enter the absolute address of the Jar file（replacing xxx.jar） and enclose it in double quotes, this is especially useful when the software prompts you that the Excel file cannot be found.

### **Options:**
`	-p <path>  `       Path to the excel file.

`	-s <sheet name> `  Name of the excel sheet.

`	-c <class name>  ` Set sorting categories, which default in 源IP.

`	-f <y/n>   `        Collapse items with the same sorting category.

`	-o <y/n>   `       Omit the same information after folding.
