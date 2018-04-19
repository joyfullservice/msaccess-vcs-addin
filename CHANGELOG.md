Change log
==========

This change log isn't being kept up to date any more, please see the git history for changes.

Version 1.0.1_beta - 7 April 2018
----------------------------
lastlink:
* started a custom ui ribbon for nondevelopers to update forms that ignores tables with extendibility through the other button
* added optional parameter for importing & exporting to ignore tables
* added a message box to display version to users on `loadVCS`, shouldn't be version of this project, but of the user's actual access database forms so for this should always be 0.01

Version 1.0.0 - 11 Mar 2015
----------------------------
jwbrookes:
* Added support for Table Data Macros
* Added support for Linked Tables (supports relative paths for linked files)
* Added support for Print Variables in Reports (Page size and orientation) 
* Added support for Relation for all types of table
* LoadVCS warnings removed when no object delete is required
* Removed elements from Report export that change constantly but don't affect import
* Fixed query import bug (complex queries being rearranged on import)
* Fixed missing constraints in table export

prmills:
* Added support for bit fields in table Import/Export
* Added support for References without GUIDs
* Export all table data with `INCLUDE_TABLES = "*"`

Version 0.12.0 - 28 Jan 2015
----------------------------
jwbrookes:
* Refactored AppCodeImportExport into several modules
* Added VCS_Loader, a module to import multiple vba modules into a database 


Version 0.11.1 - 14 Jan 2015
---------------------------
jwbrookes:
* Fixed bug in ExportTableDef function
* Removed redundant DeleteFile function (had been left commented out in the module)


Version 0.11 - 01 May 2014
-------------------------
matonb:
* Removed DeleteFile function and replaced calls to Kill with FileSystemObject.DeleteFile


Version 0.10 - 09 Mar 2014
-------------------------
matonb:
* Added DoEvents in loops to avoid "Unresponsive" state.


Version 0.9 - 15 Feb 2014
-------------------------
matonb:
*  Aggressive Sanitise, moved BaseInfo from "Block" regex to Line level.
*  Changed line level skipping to include lines with deeper indendation the follow.
  This catches split lines mostly found in BaseInfo exports.


Version 0.8 - 14 Feb 2014
-------------------------
matonb:
*  Aggressive Sanitise now excludes "BaseInfo" lines.
  These lines were seen to be randomly switching between being empty,
  not present or containing SQL on an arbitary basis.


Version 0.7 - 06 Jul 2013
-------------------------
matonb:
*  Replaced TempFile function.
*  Temporary file names now generated via external MS libraries.
*  Functions using TempFile updated to only call TempFile function once.
   *  Temporary file path and name stored in tmepFileName variable.
   *  Temporary files deleted when done.
*  Changed db declaration in ImportProject to DAO.database.


Version 0.6 - 06 Jul 2013
-------------------------

matonb:

*  AppcodeImportExport excluded from ExportAllSource
*  Added ImportProject sub-routine,  
   Deletes all forms, macros, modules and queries before calling ImportAllSource.  
   By clearing out the existing objects, you know that your database only contains  
   code from your version control database.  
   Excludes *AppCodeImportExport*

Version 0.5 - 29 May 2013
--------------------------

matonb:

*  All "exclusion" patterns are now matched by regex.
*  Added StripPublishOption constant.  
   If set to _True_ the following lines are also excluded from the export files
  * dbByte "PublishToWeb" ="1"
  * PublishOption =1
*  Added DeleteFile(FileName) function  
   The function tries to delete _FileName_ three (3) times before giving up.  
   A delay of 100ms is introduced between delete attempts should the first fail.

Version 0.4 - 19 Apr 2013
--------------------------

matonb:

*  Added dbLongBinary "DOL" to aggressive sanitize, these statements were
   appearing in queries and being flagged by git as modified in files that
   hadn't been touched by developers.

Version 0.3.2 - 8 Apr 2013
--------------------------

matonb:
* 0.3.1 Patched - Serious Problem:  SanitizeTextFiles If logic removed all
        lines containing "Begin".
* 0.3.2 Replaced if block for skipping code sections in SanitizeTextFiles with
        regular expression.

Version 0.3 - 6 Apr 2013
------------------------

bkidwell:
* Sanitize query exports.
* Fixed SERIOUS TYPO in UCS2-to-UTF-8 conversion (wrong threshold for 2 byte versus 3 byte symbol in output stream).
* AggressiveSanitize default True.

matonb:
* Added AggressiveSanitize constant, it's a number to allow for different levels in the future. ~~Default False.~~
* Added Skipping for GUID & Namemap in aggressive sanitize mode.
* ~~If AggressiveSanitize is on, also sanitize query exports.~~
* Append Number of objects imported/exported to information lines in immediate window.
* Updated readme (removed references to terminal window).
* Close all open forms and reports when importing and exporting because you can't import an open form or report.

Version 0.2 - 4 Apr 2013
------------------------

matonb:
* Added dbLongBinary "DOL" to SkipList in SanitizeTextFiles.
* Added Source directory check to ImportAllSource, pops up a message box if missing.
* Only create source directories if there is something to export.

bkidwell:
* Removed external executable for converting UCS-2-little-endian to and from UTF-8; replaced with VB6 methods.
* Added demo database to the repository.
* Removed the need for a special "export_[name]" query to export and import a lookup table.
* Added check to determine if Queries, Forms, etc. are exported from THIS database (depending on which version of Access created it) uses UCS-2-little-endian, or a legacy 8-bit Windows character set. Skip converting to/from UTF-8 if not using UCS-2, because the point of the conversion was to avoid writing 0x00 bytes in the text files and confuse diff/merge tools.

Version 0.1 - 22 Oct 2012
-------------------------

Initial release
