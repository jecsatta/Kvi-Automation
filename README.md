# Kvisthor
Little Program to automate simple tasks on Microsoft Windows
There are two configuration files: commands.txt and config.txt.

The commands.txt file contains the name of each task:
```
#this is a simple command
COMMAND:CMD
cmd.exe

COMMAND:Close Google Chrome
taskkill /im chrome.exe /f

#This is a separator
COMMAND:-

#this is a multiple line command example
#COMMAND:Restart Service
#net stop myservice
#net start myservice
```

The config.txt file contains some app configs:
```
APP_LANG:en-us
COMMAND_DBL_CLICK:taskkill /im program.exe /f
COMMAND_DBL_CLICK:taskkill /im program2.exe /f
```
Limitations: Only Command prompt commands are supported.
