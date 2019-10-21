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
 + Isolation - Manage file locking of the Excel store
 + Durability - Don't mess up the file
 
## Excel Schema

Your Excel workbook should have one spreadsheet per table. Each spreadsheet should be named to match the entity you want to administer (i.e. rename 'Sheet1/2/3' to 'Customers' if you want to have a '/Customers/' endpoint in CellServe)

Data types will be controlled by Excel's cell data type mechanism, we won't implement our own schema (for now?)

Your sheet must include headers. We assume that the **first row** is a header row and use this to set the field names.

The **first column** must always have a value, like an identifying field. You don't have to do numerical IDs, it could be some other uniquely identifying information.

## Design

**Today's design** - Process all adds/deletes immediately, `lock` on access to the workbook until avaiable

**Future design** - Add all add/delete requests to an in-memory queue, process the queue whenever a request comes in and the queue is not already being processeed on another thread.

### Add (today)

 0. An Add request comes in
 0. Enter a synchronised block for r/w the Excel book
 0. Write to the book and exit the sync block

### Add (future)

 0. An Add request comes in
 0. Add it to the in-memory cache
 0. Attempt to run the thread-safe queue digestion
 0. The queue digestion processes all Adds