# MEMDBTest
Exporter for Quicken Medical Expense Manager databases to Access format. 

## Usage
Drop the executable into the install directory of QMEM. Then drag and drop the MEM database file on top of the executable. If the database is password-protected, a window prompting for the password will appear (same UI as QMEM). Afterward, the data will start converting - this may take a while. The resulting Access database will be stored with the same name as the original + `.accdb`.

## Compiling
### Requirements
- Visual Studio 2022
- Working installation of QMEM
- .NET 4.5-4.8 SDK

### Compile
Simply open the solution and run the project.