![alt text](https://i.imgsafe.org/52/5236bd6547.jpeg "IronScripter Logo")

# Get-ITReport.PS1
### SYNOPSIS
Pulls Reports for Various things.
### DESCRIPTION
This module collects various information, specified by the argument, into CSV reports.
v1.1 Changes:
Added ability to specify one server.
Created By James A. Arnett
##### UpTime: 
*This will generate a report for the last time the server was rebooted.*
##### QFE: 
*This will generate a report of all patches and hotfixes installed based on a given number of days.*
##### ShutdownLog: 
*This will generate a report of users that have initiated a shutdown on a server based on a given number of days.*
##### Service:
*This will generate a report of services that are installed on the server.*
##### Server:
*This parameter must be used in conjunction with one or more of the above listed parameters.
This is used to specify a specific server. For usage syntax see Examples section.*
    
### REQUIREMENTS
*Powershell window must be ran as an Elevated User with access to administrative rights.*

### EXAMPLE
#### This is the basic example of the syntax

##### This will run all the above listed reports on the given list of servers.
```powershell
C:\PS> Get-ITReport
```
##### This will generate a report of all patches and hotfixes installed based on a given number of days.
```powershell
C:\PS> Get-ITReport -QFE
```
##### This will generate a report of users that have initiated a shutdown on a server based on a given number of days.
```powershell
C:\PS> Get-ITReport -ShutdownLog
```
##### This will generate a report for the last time the server was rebooted.
```powershell
C:\PS> Get-ITReport -UpTime
```
##### This will generate a report of services that installed on the server.
```powershell
C:\PS> Get-ITReport -Service
```
##### This parameter allows you to specify a specific server instead of the module prompting you for a list. This can be used with any of the parameters above.
```powershell
C:\PS> Get-ITReport -Server <Server Name>
```

### LINKS
[The Bloggin Techie](http://bloggintechie.blogspot.com/ "James' Blog")

[Chromebook Paradise](https://chromebookparadise.wordpress.com/ "Chromebook Paradise")