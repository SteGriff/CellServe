# CellServe

What if I want to use an Excel workbook as a single source of truth, but I want the kind of deliberate User Interface that a database or web application has?

CellServe is a project to span this gap by providing a REST API onto entities stored in an Excel workbook.

You can then design input forms and analytics with your choice of platform and tech.

## Advantages

 0. Easily understood storage format
 0. On-premises data - no cloud
 0. No need to install a database like SQL Server
 0. Build any input/output you want around the CellServe API
 
## Design goals

We have an aspiration to be as ACID-compliant as possible:

 + Atomicity - Commit a row to a sheet in an 'all or nothing' way
 + Consistency - Fail on writing an invalid row (based on Excel's internal validation)
 + Isolation - Manage file locking of the Excel store using an internal queue
 + Durability - Don't mess up the file
 
## Excel Schema

Your Excel workbook should have one spreadsheet per table. Each spreadsheet should be named to match the entity you want to administer (i.e. rename 'Sheet1/2/3' to 'Customers' if you want to have a '/Customers/' endpoint in CellServe)

Data types will be controlled by Excel's cell data type mechanism, we won't implement our own schema (for now?)
