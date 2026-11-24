# CellServe

What if I want to use an Excel workbook as a single source of truth, but I want the kind of deliberate User Interface that a database or web application has?

CellServe is a project to span this gap by providing a REST API onto entities stored in an Excel workbook.

You can then design input forms and analytics with your choice of platform and tech.


## Advantages

 0. Easily understood storage format
 0. On-premises data - no cloud
 0. No need to install a database like SQL Server
 0. Build any input/output you want around the CellServe API
 

## Features

 + [x] Write to an Excel sheet by POSTing
 + [x] Read all from an Excel sheet
 + [x] Read filtered data from an Excel sheet based on exact match on specified filter fields
 + [x] Get autocomplete entries for a given column
 + [x] Get schema information of the connected workbook
 + [ ] Fuzzy filters
 + [ ] Delete a row by POSTing or DELETE-ing to an existing row number
 + [ ] Update a row by POSTing to an existing row number


## Setup

The canonical solution is `src\CellServe\CellServe.slnx`. It targets .NET 10 and uses ASP.NET Core MVC. Configure the workbook path using the `Workbook:FilePath` setting in `src\CellServe\CellServe.Web\appsettings.json`, or override it with the `Workbook__FilePath` environment variable. The IIS application-pool identity must have read/write access to the workbook and its containing directory.

Restore, build, test, and publish from the repository root:

```
dotnet restore src\CellServe\CellServe.slnx
dotnet build src\CellServe\CellServe.slnx
dotnet test src\CellServe\CellServe.slnx
dotnet publish src\CellServe\CellServe.Web\CellServe.Web.csproj -c Release
```

Deploy the published output to IIS with the .NET 10 Hosting Bundle installed. CellServe does not enable cross-origin requests by default; configure approved origins in the application settings when browser clients on other origins need API access.

### Security

We expect you to run CellServe on a private local network. Any website on the network which can reach the CellServe server will be able to use its API.

Configure specific allowed origins through `Cors:AllowedOrigins` in `appsettings.json` or `Cors__AllowedOrigins__0` in the environment. Avoid using `*` for deployments that handle sensitive workbook data.


## API

Your requests should use `application/x-www-form-urlencoded`, the standard content-type for HTML forms.

Responses are `application/json; charset=utf-8`.

### `POST /Data/{Table}`

Creates a record in the specified table. Remember to set a `name` attribute on each HTML field which you want to send through. Each of these fields are expected to exist as headers in your Excel spreadsheet.

For example, if your workbook has a sheet called 'Customers', with a heading labelled 'Surname', the following form will add a Customer row, setting only the Surname field:

````
<form action="/Data/Customers" method="POST">
  <label for="Surname">Surname</label>
  <input type="text"
         id="Surname"
         name="Surname"
         placeholder="Surname"
         value="Jones">
  <button type="submit">Submit</button>
</form>
````

The response will be HTTP 201 with a confirmation of the data and operation you sent.

In case of an error, the API returns HTTP 400 with a `message` property, for example: `{"message":"No such table as Cstomers"}`


### `GET /Data/{Table}`

Gets records from the specified table. Any field that you pass as a form value will be used as a filter for an *exact match*. Without filters, the call will get everything from that sheet.

This form would get all values from the `Products` worksheet:

````
<form action="/Data/Products" method="GET">
  <button type="submit">Get all</button>
</form>
````

This form would get Products where the SKU column contains the specified value:

````
<form action="/Data/Products" method="GET">
  <label for="SKU">SKU</label>
  <input type="text"
         id="SKU"
         name="SKU"
         placeholder="SKU"
         value="ABC102">
  <button type="submit">Search</button>
</form>
````

The response will be HTTP 200 with a structured JSON representation of your data. Here's the no-filter response:

```
{
  "table": "Products",
  "operation": "Read",
  "filter": {},
  "results": [
    {
      "$ExcelRow": "2",
      "sku": "ABC101",
      "name": "Rice cake",
      "qty": "10"
    },
    {
      "$ExcelRow": "3",
      "sku": "ABC102",
      "name": "Rice flour",
      "qty": "10"
    }
  ]
}
```

N.b. We add an `$ExcelRow` property to help you find data in the source spreadsheet, and for use in Update and Delete operations.


### `GET /Data/{Table}/Suggestions?{Column}={PartialValue}`

Gets suggestions (autocomplete values) for a specified column of a specified table. The filter is not case sensitive.

This works like the GET data endpoint, but you should only specify one field-value pair.

For example, to search the Customers table for Names containing 'test', request `/Data/Customers/Suggestions?Name=test` to get a response like this:

````
{
  "table": "Customers",
  "operation": "Suggestions",
  "filter": {
    "name": "Test"
  },
  "results": [
    "Mr Test",
    "Test 1"
  ]
}
````

Some rules:

 * Your search term must be three or more characters.
 * A maximum of 20 results are returned.
 * If you don't pass a filter, you get a 400 error asking for one column/search-term value.
 * If you pass more than one filter, you get a 400 error asking for exactly one field.


## Excel Schema

An example workbook is in the 'examples' directory of this repo.

Your Excel workbook should have one spreadsheet per table. Each spreadsheet should be named to match the entity you want to administer (i.e. rename 'Sheet1/2/3' to 'Customers' if you want to have a '/Customers/' endpoint in CellServe)

Data types will be controlled by Excel's cell data type mechanism, we won't implement our own schema (for now?)

Your sheet must include headers. We assume that the **first row** is a header row and use this to set the field names.

The **first column** must always have a value, like an identifying field. You don't have to do numerical IDs, it could be some other uniquely identifying information.

We are case insensitive with:

 * Sheet names
 * Header/Field names
 * Filter values


## Design goals

We have an aspiration to be as ACID-compliant as possible:

 + Atomicity - Commit a row to a sheet in an 'all or nothing' way
 + Consistency - Fail on writing an invalid row (based on Excel's internal validation)
 + Isolation - Manage file locking of the Excel store
 + Durability - Don't mess up the file
 

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
 
## License

See LICENSE file.

If you require an enterprise license or support for this software, please contact me using `github at stegriff co uk`
