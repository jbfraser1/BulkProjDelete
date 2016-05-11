# BulkProjDelete
##Project Description

A command line .exe that allows the scripted deletion of projects from a Project Server 2007 installation. Useful for bulk project
deletes required for retention compliance. Currently, the code requires a plain text file as input for project names.

###Usage

    BulkProjectDelete -url http[s]://PWAServer/pwa/ -inputfile path\filename [-deletewsssites] [-deletearchived] [-wait] [-verify] 

###Options:

`-url pwaurl` -- Specify the url for the PWA instance on which to delete sites. Required.

`-inputfile path` -- Specify a text file listing projects to be deleted. Each project should be on a separate line. Required.

`-deletewsssites` -- WSS sites related to the deleted projects will be deleted as well. Ignored if -deletearchived is used.

`-deletearchived` -- The projects are deleted from the archive database. If not present, projects are deleted from the draft
and published databases. To delete all versions of a project, run the command twice, once with this switch and once without.

`-keeplatest` -- When deleting projects from the archive database, a project will not be deleted if it is the most recent
version of that project. Ignored if -deletearchive is not used. (New in version 1.1)

`-wait` -- Execution will pause until Project Server processes each job. This switch can be used to minimize impact on a
production server.

`verify` -- Command will not actually delete projects or WSS sites.

###Example:
     bulkprojectdelete -url https://server/pwa/ -file c:\temp\oldprojects.txt -deletewsssites

###When To Use This Tool?

Project Server 2007 includes an interface that allows the deletion of multiple projects at once (Server Settings -> Delete
Enterprise Objects.) This tool has some features lacking in the built-in interface:
 - Accuracy - the text input file allows for verification of the projects to be deleted.
 - Throttling - with the "wait" option, the project deletion will only use one thread of the queue, and will allow other
 queue jobs to be processed.
 - Automation - this tool allows for easier automation by system administrators, interfacing with a text file which can
 be generated from various sources, such as a SQL DDTS package.

###Technical Details

This code:
 1. Reads in a list of projects from the server.
 2. Reads the input file line by line. (Each project must be on its own line.)
 3. For each line, find the project in the list from the server. Store the ProjectUID.
 4. Display number of projects to be deleted, request confirmation.
 5. For each ProjectUID, place a command in the Project Server queue to delete the project. If the `-wait` option is selected,
 the program will wait until the deletion has happened before attempting to delete another project.


###Requirements

This application requires .NET Framework 3.5.

It also requires the `Microsoft.Office.Project.Server.Library.dll`. If you are running the .exe on a Project Server, it will
have access to this. If you need to run it on a client machine, you should copy the `Microsoft.Office.Project.Server.Library.dll` file from the Project Server to the same directory as the BulkProjectDelete.exe file. The .dll can be found on the server at: `C:\Program Files\Microsoft Office Servers\12.0\Bin\`, depending on where you installed Project Server.

The .exe uses the identity of whomever is logged in and executing it to access Project Server. This identity should be a
valid Project Server user with permissions to delete the appropriate projects.

###Primary Author

James Fraser
