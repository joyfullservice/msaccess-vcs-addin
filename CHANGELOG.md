Change log
==========

version 0.2 - 4 Apr 2013
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

version 0.1 - 22 Oct 2012
-------------------------

Initial release