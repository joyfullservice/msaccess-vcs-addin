msaccess-vcs-integration
========================

Synchronize your Access Forms, Macros, Modules, Queries, and Reports with a version control system.

This tool provides functions to export and import all of your Access objects (except Tables) to plain text files so that you can use a version control tool to track changes in these objects. For Access objects which are normally exported in UCS-2-little-endian encoding are automatically converted to and from UTF-8 encoding during export/import; this is to ensure that you don't have trouble branching, merging, and comparing in tools such as Mercurial which treat any file containing 0x00 bytes as a non-mergeable binary file.

This document shows how to synchronize all application code objects from an Access application with a source control system such as Kiln or Git. (It is agnostic about the actual source control systemn, but Abt IT uses Kiln which is a commercial distribution of Mercurial.)

All Forms, Macros, Modules, Queries, and Reports are synchronized. Tables and table schemas are not.

For the purposes of these instructions, assume your database is called `Application.accdb` and it is stored in `~/MyProject`.

Installing the Integration Scripts
----------------------------------

1. Load `AppCodeImportexport.bas` into a new module in your database with that exact name.
2. Copy the scripts folder as a child of `~/MyProject` .

First Commit to Your Source Control System
------------------------------------------

1. Create a repository in your `~/MyProject` folder.
2. Compact and Repair `Application.accdb` and zip it to `Application.zip` using the Send to Compressed Folder command in Windows Explorer.
3. Using your repository's tools, set the repository to ignore any `accdb` and `laccdb` files, and then add and commit the zipped Access binary file `Application.zip` as well as the `scripts` folder. Use a commit message like "Initial commit of [name] at version [number]."
4. Open the application, hit CTRL-G, and run the following VB code in the Immediate window: "`ExportAllSource`". Dismiss the popup command prompt window when it's done, and wait for the Immediate window to say the export job is "Done."
5. Using your repository's tools, add and commit all the new files that were created in the source folder under `~/MyProject` . Use a commit message like "Initial commit of all source code for [name] at version [number]".
6. Publish your repository to your preferred central sharing location.

Committing New Progress and Pulling Changes from Other Developers
-----------------------------------------------------------------

1. Open the application, hit CTRL-G, and run the following VB code in the Immediate window: "`ExportAllSource`". Dismiss the popup command prompt window when it's done, and wait for the Immediate window to say the export job is "Done."
2. Using your repository's tools, commit all the new files that were created in the source folder under `~/MyProject` . Use an appropriate commit message to describe your changes.
3. Pull new upstream changes (if any exist) from your central sharing location used by all developers. If necessary address any merge conflicts using your repository's merge and conflict resolution tools. If any work was done in this step, commit these changes to your local repository as well.
4. Push all local and merged changes back to the central sharing location.
5. Go back into the Access Immediate window (CTRL-G) and run the following VB code: "`ImportAllSource`". Dismiss the popup command prompt window when it's done, and wait for the Immediate window to say the export job is "Done."

Committing a New "Release" of Your Project
------------------------------------------

1. There may be application changes that aren't covered in the source code for Forms, Macros, Modules, Queries, and Reports. To make sure these changes are recorded, Compact and Repair `Application.accdb` and zip it to `Application.zip` (replacing the old copy) using the Send to Compressed Folder command in Windows Explorer. Commit the new Application.zip to your repository with a commit message like "Full application binary for release [number]".
2. Follow the usual steps in the previous section "Committing New Progress".
3. Use your repository's "tag" function to tag your last commit with the release number/name.

Caveats
-------

* If you make any changes to the script used in this process, the `AppCodeImportExport` module, they will not be automatically imported when any developer runs the ImportAllSource method. The code skips this file because it causes a conflict when trying to update a module that is actively being executed.
* The import and export code does not handle deleted objects. When you notice that a developer upstream has pushed a change that deletes an Access object, you must manually delete that object in your own copy of the database file, and be sure it's not exported again and re-added to the repository.
