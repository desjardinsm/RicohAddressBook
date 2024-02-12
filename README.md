# RicohAddressBook

[![Build status](https://ci.appveyor.com/api/projects/status/6wq08909v4c6cbjn?svg=true)](https://ci.appveyor.com/project/desjardinsm/ricohaddressbook)

A PowerShell module to manage Ricoh printer address books.

Based on the PowerShell module created by Alexander Krause.

## Exports

-   `Get-AddressBookEntry`
-   `Add-AddressBookEntry`
-   `Update-AddressBookEntry`
-   `Remove-AddressBookEntry`
-   `Get-Title1Tag` (this is only useful for getting the proper Title1 for a
    letter, to be used in the Add- or Update- functions)

## Installation

Download the files from the [latest release][latest_release] (or download them
from the "Module" directory of the repository itself) and place them in a
directory named `RicohAddressBook` somewhere in the `$env:PSModulePath`.

[latest_release]: https://github.com/desjardinsm/RicohAddressBook/releases/latest

For a user-installation, this would typically be something like
`%USERPROFILE%\Documents\WindowsPowerShell\Modules`.

Alternatively, you can place the files in an accessible location and run the
following command manually when you wish to use it:

```powershell
Import-Module "<path to module>\RicohAddressBook.psd1"
```

## Credits

This would not have been possible without the work done in the following projects:

-   [Alexander Krause's PowerShell module][ps_module_archive] (which was hosted
    on the TechNet gallery)
-   [Ricoh.NET](https://github.com/gheeres/Ricoh.NET)
-   [libmfd](https://github.com/adam-nielsen/libmfd)

[ps_module_archive]: https://web.archive.org/web/20200318044655/https://gallery.technet.microsoft.com/scriptcenter/Ricoh-Multi-Function-27aeea71
