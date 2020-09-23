Attribute VB_Name = "MResourse"
'MSHFlexGrid

'Button resource
Public Const rStartMonitoring = "Start Monitoring"
Public Const rStopMonitoring = "Stop Monitoring"

'Datagrid resource
Public Const rTitle = "Title"
Public Const rAddress = "Address"
Public Const rCategory = "Category"
Public Const rUserName = "User Name"
Public Const rPassword = "Password"
Public Const rFirstDate = "First Visit Date"
Public Const rLastDate = "Last Visit Date"
Public Const rVisitCount = "Total Visit"

'StatusBar resource
Public Const rTime = "Time"
Public Const rToday = "Today"
Public Const rTotal = "Total"
Public Const rRecord = "records"

'Images & Icons
Public Const rSysTrayIcon = "images\monitor.ico"
Public Const rSysTrayDisIcon = "images\unmonitor.ico"


'SSTab


'File Paths
Public Const con_INI_File As String = "linkmonitor.ini"
Public Const con_DB_Name As String = "InternetAddresses.mdb"


'Encryption key
Public Const ENCRYPT_KEY As String = "lm"


'HTML Index
Public Const htmlHeader As String = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'><html><head><title>$DocumentTitle$</title><meta http-equiv='Content-Type' content='text/html; charset=iso-8859-9'><style type='text/css'> div.code { overflow: auto; padding: 5px; /* was 3px */ color: #000066; background-color: #F0F0F0; font-family: verdana; /* added */ font-size: 10px; border: 1px solid blue; } tr.header { font-size: 11px; font-weight: bold; background-color: #06E4BD;} tr.first { background-color: #ADCADC;} tr.second { background-color: #BDDAEC;} </style> <script language='JavaScript'>var initialBackColor;</script> </head><body><div class='code'><table border='0'><th><tr class='header'> <td>Number</td> <td> Title</td> <td> Address</td> <td> Category</td> <td> User Name</td> <td> Password</td> <td> First Visit Date</td> <td> Last Visit Date</td> <td> Total Visit</td></tr></th>"
Private Const mouseEfect As String = "onMouseOver='initialBackColor = this.style.backgroundColor; this.style.backgroundColor = ""#FFCCCC"";' onMouseOut='this.style.backgroundColor = initialBackColor;'"
Public Const htmlTableRowTemplateFirst As String = "<tr class='first' " & mouseEfect & "> <td>$Number$</td> <td> $Title$</td> <td> <a href='$Address$' target='_blank'>$Address$</a></td> <td> $Category$</td> <td> $UserName$</td> <td> $Password$</td> <td> $FirstVisitDate$</td> <td> $LastVisitDate$</td> <td> $TotalVisit$</td></tr>"
Public Const htmlTableRowTemplateSecond As String = "<tr class='second' " & mouseEfect & "> <td>$Number$</td> <td> $Title$</td> <td> <a href='$Address$' target='_blank'>$Address$</a></td> <td> $Category$</td> <td> $UserName$</td> <td> $Password$</td> <td> $FirstVisitDate$</td> <td> $LastVisitDate$</td> <td> $TotalVisit$</td></tr>"
Public Const htmlFooter As String = "</table></div></body></html>"






