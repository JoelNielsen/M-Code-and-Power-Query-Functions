A collection of code used in an Enrolments environment for a school to validate large datasets of students/customers.

Most code is specific for Excel Online and work within cells or Power Query as the documents are live sharepoint files.
Excel Online does not currently support Data Models or VBA code, so that is why a lot of the code would look odd and could be done much easier with in-built tools through Data Models.

Another limitation of Excel Online is that Pivot Tables referencing multiple tables breaks/does not work; which is why most comparison functions are done on separate sheets to reduce the need for the user to understand what is going on in the background to compare tables.

Please note: The conversion of PDF to excel table is incredibly taxing on most department-issued computers and will run CPU at 100% for many hours until it's converted. Probably don't do that.
