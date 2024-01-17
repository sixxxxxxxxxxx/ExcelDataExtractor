![A nuget package used for reading from and writing to excels.](https://raw.githubusercontent.com/sixxxxxxxxxxx/ExcelDataExtractor/main/excel.jpg)

# ExcelDataExtractor

A nuget package used for reading from and writing to excels.

## Badges

[![MIT License](https://img.shields.io/badge/License-MIT-green.svg)](https://choosealicense.com/licenses/mit/)

stable release version: ![version](https://img.shields.io/badge/version-1.0.30-blue)

## Tech Stack

**C#, .Net6.0, .NetStandard2.1**, 


## How Do I Get Started

First, install NuGet. Then, install ExcelDataExtractor from the package manager console:

```C#   
   NuGet\Install-Package ExcelDataExtractor -Version 1.0.30
```
 This command is intended to be used within the Package Manager Console in Visual Studio, as it uses the NuGet module's version of Install-Package.


Or from the .NET CLI as:
```C#   
   dotnet add package ExcelDataExtractor --version 1.0.30
```

Finally, import into the file:
```C#   
   using ExcelDataExtractor;
   using ExcelDataExtractor.Dtos.Responses;
```

## Features

- Write to excel
- Read from excel

## Sample usage

```C#
   Excel.WriteToExcel(collection: List<T>(), worksheetName: string, unlockedColumns: int[], hiddenColumns: int[]);           
```
- Input

| Parameters	   | Type		| Description										   |
| :--------		   | :-------	| :-------------------------						   |
| `collection`,	   | `List<T>()`| **Required**. List of model to be converted to excel |
| `worksheetName`  | `string`	| **Required**. Name of file after write is completed  |
| `unlockedColumns`| `int[]`	| **Required**. Columns that should be unlocked		   |
| `hiddenColumns`  | `int[]`	| **Required**. Columns that should be hidden          |

- Output

| Type     |
| :------- |
| `ExcelWriteResponse` |

---
```C#
   Excel.ReadFromExcel<T>(excelFile, requiredHeaders: string[], nullableColumns: string[], columnsToSkip: string[], columnToCheckForDuplicates: string, uniqueColumns: string[]);
```

- Input

| Parameters	              | Type	       | Description										   |
| :--------		              | :-------       | :-------------------------						       |
| `excelFile`,	              | `IFormFile`    | **Required**. Excel to be converted to object         |
| `requiredHeaders`			  | `string[]?`	   | **nullable**. Columns that cannot be empty            |
| `nullableColumns`           | `string[]?`	   | **nullable**. Columns that can be empty		       |
| `columnsToSkip`             | `string[]?`	   | **nullable**. Columns that should be skipped          |
| `columnToCheckForDuplicates`| `string`	   | **Required**. Columns that duplicates are not allowed |
| `uniqueColumns`             | `string[]?`	   | **nullable**. Columns that must be unique             |

- Output

| Type					 |
| :-------				 |
| `ExcelReadResponse<T>` |




## Thanks to all Contributors

Maintainers:

- [sixxxxxxxxxxx](https://github.com/sixxxxxxxxxxx)
- [bubethedev](https://github.com/bubethedev)

## Contributing

Contributions are always welcome!

See `contributing.md` for ways to get started.

Please adhere to this project's `code of conduct`.
