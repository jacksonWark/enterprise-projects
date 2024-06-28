# Enterprise Projects

Contained in this repository is a selection of projects I completed for my enterprise clients as a freelance developer and consultant. All of these project I contributed > 50% of the design, implementation, testing and support. 5 of the 6 projects are source code in Delphi from varying versions of a RADStudio environment, while 1 was completed in Visual Studio 2019 .NET Windows Forms. 


## Projects I was 90-100% Responsible for: 

### Crystal Reports .NET Compatibility Module Project (CryRptsDotNet/ folder)
The client's ERP uses Crystal Reports extensively, as do all their custom applications to leverage common technology and maintain a consistent experience for their employees. With newer versions of the ERP were newer versions of Crystal Reports, which had dropped support for ActiveX in favor of .NET and therefore support for Delphi development directly. Using a 3rd party library for cross-language modules (Remobjects Hydra), I created a DLL module implementing common use cases for Crystal reports to be used in custom applications. 

### Fuel Surcharge Updater Application (FSCUpdater/ folder)
In order to reduce the amount of data entry for the client's process of maintaining many Fuel Surchage codes for clients and interliners, I designed and developed an application in Delphi to scrape commonly used reference rates, as well as update batches of records in the client's database tables to manage these Fuel Surcharge codes. 

### High Radius Credit Module Data ERP Integration (HRCxTM/ folder)
The client chose to purchase a credit analysis service from a company called HighRadius. They chose to develop their integration into their ERP system, and I took the lead on the design and implementation. This Delphi application runs as a windows service that manages a daily data exchange between a HighRadius SFTP server, and the client ERP database. This includes client account ID generation, data mapping, notifying staff of significant changes, and generating reports.

### Shared Services Migration Common Modules (Shared_Services_Packages/ folder)
The client migrated from a locally managed on-premises desktop and remote-access server environment, to a shared-services off-site environment managed by the parent organization. With the IT department we had to ensure that all of their custom appilcations were compatible with this new environment. We also took the opportunity to update, standardize, and modularize common functionalities across their custom applications, as well as make the applications generally more configurable by IT personnel. Modules include a crytpography module that encrypts sensitive information, standard configuration files, database configuration interface, and Microsoft Graph Mail API module.


## Projects I contributed between 50-90%:

### TGI Connect Trailer Tracking API ERP Integration (TGIConAPI/ folder)
Client purchased trailer-tracking GPS units and were provided with an JSON API to track the location updates.  
I helped create and ERP integration that requests & parses the returned JSON data for client’s IBM DB2 SQL DB. This required translating the GPS coordinates to zone based positions that are used by the ERP.


### TST-CF Express Accessorial API ERP Integration (eFrtAPI/ folder)
Client needed an ERP integration for a freight shipment services XM API from another freight company. I developed the code that parses/generates XML data to/from the database.

