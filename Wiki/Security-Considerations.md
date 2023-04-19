We have worked hard to make this add-in as accessible to as many environments as possible, including those that have higher security requirements. It does not require administrative privileges to install or run, and utilizes a trusted location to allow global settings to remain more restrictive.

The entire project with all source code is available on GitHub at: [joyfullservice/msaccess-vcs-integration](https://github.com/joyfullservice/msaccess-vcs-integration) See [LICENSE.txt](https://github.com/joyfullservice/msaccess-vcs-integration/blob/master/LICENSE.txt) for specific terms and disclaimers. (BSD-style license)


# Minium Security Requirements
The following settings need to be allowed in the Security Center for the add-in to be able to interact with databases and function as designed.
- **Allow Trusted Locations** (Including network locations, if the database is not on the development computer.)
- **Allow Trusted Documents** - In some environments the add-in file must be trusted as well.
- **ActiveX Settings - Prompt before enabling** - This is required for importing XML content, such as table definitions.
- **COM Add-ins** - Needed to use the Ribbon toolbar. (Primary launcher for Version 4 and above)

# Normal Behavior
This add-in was created in Microsoft Access VBA, which gives a powerful platform for developing a highly performant add-in with easily accessible source code and live debugging capabilities.

As with any Microsoft Access add-in, this powerful tool comes with inherent security considerations. This page is intended to guide you in understanding some of the fundamental operations of the add-in, and how this might affect the security of your systems and data.

## Export to Source Code
An existing database can be exported to (primarily) text-based source files that can be committed to a file-based version control system such as Git. While primarily focused on the database *component* objects like forms, queries, table definitions and VBA source code, you can optionally choose to export *data* from tables as well.

Here are a few considerations to think about when exporting to source:
- Some connection information, such as ODBC connection strings may be present in your exported source files.
- Paths to network locations and the database location may be present in some source files.
- Any table data that is exported could contain PII or other sensitive content, which could then be easily committed to version control.
- Access to the exported source files should be protected, just like you would protect the database itself.

## Build from Source Files
This add-in provides the ability to build a functioning database application entirely from the source files saved during an export. This is a very powerful collaboration tool when combined with a version control system. Changes at the source file level can be merged and used to build a specific version of a database application.

Here are some potential security considerations when building from source:

- Probably the most important aspect here is to protect your *source files* from any unauthorized modifications.
- The build operation can include "hooks" to run specific VBA code during or after the build process. This is a very powerful feature, but could be used maliciously by someone with access to the source files.
- Any source files from external locations should be carefully reviewed before including in a project for build.

## Add-in Behavior
During the export and/or build process, the add-in uses the following:
- **`CreateObject()`** - Create ActiveX objects like `Shell.Application`. (Most objects are created using early binding.)
- **Windows API Calls** - Used in many areas including performance timing, string hashing, and time zone determination.
- **Files** - Creates, modifies and deletes source files for the project.
- **Registry** - Reads and writes values to the Windows registry. (Current User Hive)

# File System Access
Most of the processing happens in the source folder for the database, but the user's Windows temporary folder is also used for processing and comparing source files.

The add-in is installed in the user's `AppData` folder, and includes the COM add-in ActiveX DLL that provides the application Ribbon interface.

# Registry Access
The registry is used to store certain program settings, as well as transient values used during some processes like building from source. Other registry settings are used to determine security settings in the current environment.

The following sections are both read and written by the add-in.

**Trusted Locations** - Used when making the add-in installation path a trusted location.

`HKCU\Software\Microsoft\Office\[version]\Access\Security\Trusted Locations\`

**Add-in Menu** - Adds shortcuts to the Microsoft Access add-ins menu. (The primary way to launch the add-in before the ribbon was added in version 4.)

`HKCU\Software\Microsoft\Office\[version]\Access\Menu Add-Ins\`

**Ribbon UI** - The ribbon was created in twinBASIC as a COM add-in (ActiveX DLL) which is self-registering using a call to `regsvr32` during the installation process. (COM registration entries are also created in the current user hive.)

`HKCU\SOFTWARE\Microsoft\Office\Access\Addins\MSAccessVCSLib.AddInRibbon`

**Local Settings** - Program settings and transient values. (VBA's native `GetSetting`/`SaveSetting` functions.)

`HKCU\SOFTWARE\VB and VBA Program Settings\MSAccessVCS\`

# Telemetry
This add-in does not communicate in any way with remote servers or systems. It does not check for updates or send any usage statistics, and does not require an Internet connection to use the tool.

That being said, it *can* interact with linked tables and potentially remote servers and systems in connection with databases being exported or built from source. This is entirely within the configuration and intended usage of the add-in.

For example, if your database has a linked table on a SQL server, the add-in will attempt to connect to this remote table to retrieve a list of columns that are written to a source file. This all happens locally within the context of your network environment and development computer.

# Encryption / Cryptography
Some countries and corporate environments have specific requirements when it comes to cryptography. Here are a few ways this add-in uses cryptography:
- **Zip** - Database theme files (`*.thmx`) natively use Zip compression.
- **File Hashes** - Used to detect and index changes to files. These hashes are `SHA256` by default, but can be changed to `SHA1` or `SHA512` in the add-in options.
- No reversible encryption is used in this project at this time.

# VBA References
In addition to the built-in Microsoft Access references, this add-in also uses the following references:
- **Microsoft ActiveX Data Objects 6.1 Library** - For ADO interactions with some database objects.
- **Windows Script Host Object Model** - For registry access.
- **Microsoft Office [version] Object Library** - For open dialogs and other general Office integration.
- **Microsoft Scripting Runtime** - For `FileSystemObject` and other related components for interacting with the file system in a Unicode compliant way.
- **Microsoft Visual Basic For Applications Extensibility** - For interaction with VBA source code components.
- **Microsoft VBScript Regular Expressions 5.5** - For parsing source file content.
- **Microsoft XML, v6.0** - For managing XML source file content.
- **Microsoft Forms 2.0 Object Library** - For interacting with native VBA forms (different from Microsoft Access forms).
- **UIAutomationClient** - For interactions with navigation pane objects, such as selecting a single object for export.

# Additions / Updates
If you have further items that would be helpful to add to this page, please feel free to open an issue or create a pull request!

Reviewed 4/19/2023 by @joyfullservice