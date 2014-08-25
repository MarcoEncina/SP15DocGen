SP15DocGen
==========

SPDocGen - SharePoint 2013 Documentation Generator

Upgraded version of [SPDocGen - SharePoint Documentation Generator](https://spdocgen.codeplex.com/) by [SharePointSezai](https://www.codeplex.com/site/users/view/SharePointSezai) which works for MOSS 2007.
An upgraded version for SharePoint 2010 can be found [here](https://lekman.codeplex.com/releases/view/55953) by [Tobias Lekman](https://www.codeplex.com/site/users/view/Lekman).

You can download the executables [here](https://github.com/siaf/SP15DocGen/tree/master/Binary)

## Project Description#
Automatically generate SharePoint Farm Documentation in a few seconds.

The application uses the SharePoint Object Model API to produce an XML file containing data from a SharePoint Farm Deployment. The XML file is then transformed into a .doc file (WordML) via XSLT. 

1. Unzip to a folder on a WSS 3.0 or MOSS 2007 server.
2. Logon to the server as a Farm Administrator with permissions to write to the folder that Sezai.SPDocGen.Console.exe runs from.
3. Produces an XML and DOC file as output, modify the XSLT file as you like.

Full Source Code for the solution is provided so you can tweak it as needed.


## What does this tool document?
* Servers in the farm
* Server Services
* Databases
* Database Name, Server and Disk Size Required for Backup
* Farm Services
* Service Job Definitions for each Farm Service
* Service Job Definition Name, Schedule, Last Time Run
* Web Applications
* Name, ID, Site Collections, Content Databases and Application Pools
* Site Collections
* Title, URL, Web Application, Web Count, Users Count, Disk Space Used, Owners and Site Collection Administrators
* Farm Application Pools
* Name, Identity and Password
* Farm Solutions
* Name, ID, Deployment Status, Last Deployment Action Time, Last Operation Details, Deployed Servers, Deployed Web Applications
* Farm Feature Definitions
* Feature Name, ID, Scope, Root Directory, Solution ID, Receiver Assembly and Class


If the executable throws an error, copy out the exception message and post it to the Issue Tracker for this project.

