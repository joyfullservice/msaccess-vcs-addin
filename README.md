msaccess-vcs-integration
========================

[![Join the chat at https://gitter.im/timabell/msaccess-vcs-integration](https://badges.gitter.im/timabell/msaccess-vcs-integration.svg)](https://gitter.im/timabell/msaccess-vcs-integration?utm_source=badge&utm_medium=badge&utm_campaign=pr-badge&utm_content=badge)

About
-----

Synchronize your Microsoft Access Database definition with a version control system.

Included in the export/import:

* Queries
* Forms
* Reports
* Macros
* Modules
* Table Data
* Table Definitions
* Table Data Macros

Not included in the export/import:

* Customized toolbars, toolbar items, and menu items
* Any external files
* Pretty much anything that is not accessible by browsing the design, properties, and code of a Query, Form, Report, Macro, or Module object.

*This code is built and tested on Microsoft Access 2010/2013. It will probably work in earlier/later versions, but YMMV.*

This README shows how to synchronize all application code objects from an Access application with a source control system such as Mercurial or Git. (The provided import/export module is agnostic about the actual source control system you use.)

Encoding
--------
For Access objects which are normally exported in `UCS-2-little-endian` encoding , the included module automatically converts to the source code to and from `UTF-8` encoding during export/import; this is to ensure that you don't have trouble branching, merging, and comparing in tools such as Mercurial which [treat any file containing 0x00 bytes as a non-diffable binary file](https://www.mercurial-scm.org/wiki/BinaryFiles).

Output
------
The module will put the files in a folder called `source` within the same folder as your database file. The import expects the files to be in the same folder.


Installing the Integration Scripts
----------------------------------

For the purposes of these instructions, assume your database is called `Application.accdb`.

1. Load `VCS_Loader.bas` into a new module in your database with that exact name.
 1. Go to the VBA editor (CTRL-G) and select "File" > "Import File..."
    (or you can just drag and drop the file from windows explorer into the vba editor module list).
 2. Select the `VCS_Loader.bas` file.
 3. Save the file (CTRL-S).
2. Type "`loadVCS`" into the immediate window followed by the directory where the other VCS files are located. If you don't specify a directory then it is assumed that the VCS code is contained in a folder called 'MSAccess-VCS' in the database directory;
e.g. `loadVCS "C:\Users\MyUserAccount\Documents\Access-Proj\MSAccess-VCS\"` - the trailing slash is required
or `loadVCS` will not import the new modules.
3. Edit your `VCS_ImportExport` and change the constant `INCLUDE_TABLES` to list any lookup tables that function more as part of your application code than as client data. (For example, "Countries", "Colors", and things like that.)

Configuring export
--------------------------------

By default, no table data is exported. You must specify which tables' data to include in the export/import process by editing the `INCLUDE_TABLES` variable in the supplied module. For example you might have "Countries" or "Colors" tables that populate dropdown lists. You shouldn't include regular data tables containing actual records, because this data doesn't belong in version control.

Additionally, if a type of data should not be exported, change the "Export_" constants to `False` in `VCS_ImportExport.bas`. Report, Query, Form, Macro, Module, and Table exports can be disabled individually.

Supplied databases
------------------

In the `demo\` folder there's a blank database that you can use with to provide with your source-controlled files, or to test the import; and a demo database with a sample of all the things that this project can import/export for trying the project out and testing any code changes made to the project.

First Commit to Your Source Control System
------------------------------------------

1. Create a repository in the folder containing your database.
2. Compact and Repair `Application.accdb` and zip it to `Application.zip` using the Send to Compressed Folder command in Windows Explorer.
3. Using your repository's tools, set the repository to ignore any `.accdb` and `.laccdb` files, and then add and commit the zipped Access binary file `Application.zip`. Use a commit message like "Initial commit of [name] at version [number]."
4. Open the application, hit CTRL-G, and run the following VB code in the Immediate window: "`ExportAllSource`". Wait for the Immediate window to say the export job is "Done."
5. Using your repository's tools, add and commit all the new files that were created in the `source` folder. Use a commit message like "Initial commit of all source code for [name] at version [number]".
6. Publish your repository to your preferred central sharing location.

Committing New Progress and Pulling Changes from Other Developers
-----------------------------------------------------------------

1. Open the application, hit CTRL-G, and run the following VB code in the Immediate window: "`ExportAllSource`". Wait for the Immediate window to say the export job is "Done."
2. Using your repository's tools, commit all the new files that were created in the source folder. Use an appropriate commit message to describe your changes.
3. Pull new upstream changes (if any exist) from your central sharing location used by all developers. If necessary address any merge conflicts using your repository's merge and conflict resolution tools. If any work was done in this step, commit these changes to your local repository as well.
4. Push all local and merged changes back to the central sharing location.
5. Go back into the Access Immediate window (CTRL-G) and run the following VB code: "`ImportAllSource`". Wait for the Immediate window to say the export job is "Done."

Committing a New "Release" of Your Project
------------------------------------------

1. There may be application changes that aren't covered in the source code for Forms, Macros, Modules, Queries, and Reports. To make sure these changes are recorded, Compact and Repair `Application.accdb` and zip it to `Application.zip` (replacing the old copy) using the Send to Compressed Folder command in Windows Explorer. Commit the new `Application.zip` to your repository with a commit message like "Full application binary for release [number]".
2. Follow the usual steps in the previous section "Committing New Progress".
3. Use your repository's "tag" function to tag your last commit with the release number/name.

Loading/updating a database from the exported files
---------------------------------------------------
1. Create a new Access database (or use the supplied `demo\blank.accdb`).
2. Follow the instructions for installing the scripts.
3. Open the VBA editor (CTRL-G) and run the following VB code in the Immediate window: "`ImportProject`". You will be presented with a warning telling you that all database objects are about to be deleted, allowing you to cancel the operation if you change you mind.
4. Wait until the code finishes executing, Compact and Repair the database.

Caveats
-------
* If you make changes to or add a new module, be sure to save it in the VB Editor window or else it will not be exported.
* If you make any changes to the scripts used in this process, the `VCS_` modules, they will not be automatically imported when any developer runs the `ImportProject` method. The code skips these files because it causes a conflict when trying to update a module that is actively being executed.


Contributing
============

Pull requests, issue reports etc welcomed.

https://github.com/timabell/msaccess-vcs-integration seems to currently be the
most actively maintained branch, and [Tim Abell](https://github.com/timabell)
will generally accept pull requests to keep the project alive but has minimal
capacity to ensure correctness so please try and keep the quality as good as
you can. Thanks! Several other people have been made collaborators to help
keep the repo alive. Better ideas for running this project would be welcome
https://github.com/timabell/msaccess-vcs-integration/issues/32
