# Excel
The Excel Adapter allows customers to read and write Excel Documents

# Excel Adapter overview

The Excel Adapter allows you to read / write excel data. 


## Minimum Requirements : 

WTX v8.4.1.x <br>
ITX v9.0.0.x

## Command Alias : 

Adapter commands can be specified by using a command string on the command line or creating a command file that contains adapter commands. The execution command syntax is:

-IM[alias] card_num <br>
-OM[alias] card_num


where -IM is the Input Source Override execution command and -OM is the Output Target Override execution command, alias is the adapter alias, and card_num is the number of the input or output card. The following table shows the adapter alias and its execution command.


Adapter 	:  EXCEL <br>
Alias 	        :  EXCEL <br>
As Input        :  -**IMEXCEL**card_num <br>
As Output       :  -**OMEXCEL**card_num <br>    	  


## EXCEL Adapter commands

The following table lists valid commands for the EXCEL Adapter, the command syntax, and whether the command is supported (✓) for use with data sources, targets, or both.

Location (-F)     : Location of the Input / output Excel Document<br>
WorkSheet (-W)	  : Worksheet with in the Excel Document<br>
Template Location (-P )  : Location of the Excel Template<br>. Supported on the Target only and is optional argument

## EXCEL Adapter Command Line Examples
###Inbound Adapter : 

Command Line : -F test.xlsx -W Sheet1 <br>
GET RULE : =PARSE(GET("EXCEL", "-W Sheet1 -T", Input)) <br>

###Outbound Adapter : 

Line : -F testout.xlsx -W Sheet1 <br>
Command Line with Template : -F testout.xlsx -W Sheet1 -P template.xlsx <br>
PUT RULE : PUT("EXCEL", "-F testout.xlxs -W Sheet1 -P template.xlxs -T", Input)) <br>


## EXCEL Adapter Installation Instructions 
### Design Studio

a) Drop com.ibm.websphere.dtx.ui.importers.excel_<VRM>.jar, com.ibm.websphere.dtx.ui.importers.excel.nl_<VRM>.jar
to <WTX INSTALL>\DesignStudio\wtx_eclipse\eclipse\plugins <br>
b) Drop m4excel.jar in to <WTX INSTALL> <br>
c) Edit adapters.xml and add the following line <br>

<M4Adapter name="Microsoft Excel" alias="EXCEL" id="164" type="app" class="com/ibm/websphere/dtx/m4excel"/> <br>

d) Download Apache POI artifacts from [Apache POI](https://poi.apache.org/download.html) version 3.13. Copy commons-codec-xx.jar, commons-logging-xx.jar, log4j-xx.jar, poi.jar, poi-excelant.jar, poi-ooxml.jar, poi-ooxml-schemas.jar, xmlbeans-2.6.0.jar to <WTX INSTALL DIR> <br>

d) Invoke cleanextenderstudio.bat.
 
### Runtime

a) Drop m4excel.jar in to <WTX INSTALL> on windows and to <WTX INSTALL>/libs on UNIX <br>
b) Edit adapters.xml and add the following line (UNIX, the file is present under config folder) <br>

<M4Adapter name="Microsoft Excel" alias="EXCEL" id="164" type="app" class="com/ibm/websphere/dtx/m4excel"/> <br>

c) Download Apache POI artifacts from [Apache POI](https://poi.apache.org/download.html) version 3.13. Copy commons-codec-xx.jar, commons-logging-xx.jar, log4j-xx.jar, poi.jar, poi-excelant.jar, poi-ooxml.jar, poi-ooxml-schemas.jar, xmlbeans-2.6.0.jar to <WTX INSTALL DIR> (UNIX, copy to <WTX INSTALL>/libs folder)  <br>

