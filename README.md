# FellowshipOneCheckinAutomation
Code base for Autoit Application that automates the Fellowship One Check-in app

This is the First version of my automation app. It works on Windows 7/8/10, and is configured either by an XML file or by passing parameters thru the command line. The code is heavly commented straight forward.
The application is started with a scheduled task that runs a batch file. The batch calls the program as well as screen savers, powershell scripts and performs other tasks.

 --- edit 05/17/16 
Major changes in Sending Activity Code in 1.0.0.30 works much better. Also added email function for reporting, switched to a powershell script for automation triggered by task scheduler before service times, and special screen savers for errors.

Overall I now have verry few issues with this system. Out of 30 PC's I get maybe 2 error email's per service. Really beats having to go to 30 machines and configure all of them by hand. Most of the problems come from internet issues.
