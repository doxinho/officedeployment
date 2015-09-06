# Office 365 Deployment & Removal

Currently utilizing version 3.6.5 of the Powershell Application Deployment Toolkit.


1. Edit Download.xml
  1. Edit the SourcePath to your directory.
  2. Edit the ProductID to match your subscription. Full list of ProductIDs available [here](https://support.microsoft.com/en-us/kb/2842297).

2. Edit Installation.xml
  1. Edit the Product ID, language, logging (if desired), autoactivation, and updates. Full documentation available here.

3. Run download.bat, downloads configured setup files

4. Adjust Deploy-Application.ps1 to your desired configuration. By default, it does the following:
  1. Prompts the user that there is an application install.
  2. Forces close all Office applications & Internet Explorer, gives user 60 seconds.
  3. Removes all existing versions of Office (2003, 2007, 2010, 2013, 2013 Click-to-Run, Office 365 & associated applications) utilizing the Microsoft created vbscripts in their Fix-It tools
  4. Installs Office 365
  5. Suppresses the first run dialogs for Office 365 for all users
  6. Dialog prompt that the installation is complete.
  7. Prompts user to restart, forces within 60 seconds

5. 

You can easily adjust this script to your desired needs. It is easy to run fully silently and to adjust the messages/prompts. The PSADT documentation is excellent. 

Credit goes to the sample scripts and #sysadmin on irc.synirc.net for assisting with the development of this script.
