# FellowshipOneCheckinAutomation
Code base for Autoit Application that automates the Fellowship One Check-in app

This is the First version of my automation app. It works on Windows 7/8/10, and is configured either by an XML file or by passing parameters thru the command line. The code is heavly commented and straight forward.
The application is started with a scheduled task with multiple defined time based triggers that run a powershell script. The script calls the automation program, screen savers, options program as well as perform other task.

 --- edit 05/17/16 ---
 
Major changes in Sending Activity Code in 1.0.0.30 works much better. Also added email function for reporting, switched to a powershell script for automation triggered by task scheduler before service times, and special screen savers(blue screen) for errors.

Overall I now have very few issues with this system. Out of 30 PC's I get maybe 2 error email's per service. Really beats having to go to 30 machines and configure all of them by hand. Most of the problems come from internet issues.

*Note*

To build you will have to add the UIAWrappers.au3 file to your include folder in your autoit install folder. I've added it to the repo(uploaded version has one line commented out on line 560 to stop logfiles for every run) but here is the link if you want to get it yourself. Use version v0 51.zip: https://www.autoitscript.com/forum/topic/153520-iuiautomation-ms-framework-automate-chrome-ff-ie/
