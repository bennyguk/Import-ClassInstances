# Import-ClassInstances
A script to import class instances and their properties, relationships and file attachments for Work Item or Configuration item based classes exported using the Export-ClassInstance.ps1 script (https://github.com/bennyguk/Export-ClassInstances) in to System Center Service Manager.
    
This script could be useful if you need to make changes to a custom class that are not upgrade compatible and have been exported using the Export-ClassInstances.ps1 script (https://github.com/bennyguk/Export-ClassInstances).

## To use
The script requires SMLets to be installed (https://github.com/SMLets/SMLets/releases) as well as the cmdlets distributed with the Service Manager console.

The script has the following parameters:
* ClassName - The Name property of the class rather than DisplayName to be imported. Has to match the class name used to create the CSV files.
* FilePath - A folder path to load the CSV files and file attachments.
* ComputerName - The hostname of your SCSM management server.
* FileName - (Optional) - The name of the CSV file that was exported with Import-ClassInstances.ps1. The script expects the relationship csv file to be called *FileName*-relatrionships.csv.

## Additional Information
This script started life simply as a way of exporting all instances of a particular class from our production environment in preparation of an upgrade to a custom Configuration Item based class Management Pack, but I soon started to wonder if it would be possible to programmatically export properties and relationships of any Work Item or Configuration Item. 

It is very much a work in progress and does have limitations with some relationship types such as SLAs, Request Offerings and Billable Time user relationships. If you have an idea about how to deal with these relationship types or any other improvements, I would be delighted for you to contribute :)
