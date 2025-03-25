- [Is there a way to use a ribbon with this add-in?](#is-there-a-way-to-use-a-ribbon-with-this-add-in)
- [Why are some issues/ideas considered out of scope for this project?](#why-are-some-issuesideas-considered-out-of-scope-for-this-project)
- [Why am I seeing a large number of "changed" files after building my project from source?](#why-am-i-seeing-a-large-number-of-changed-files-after-building-my-project-from-source)
- [How do I also export data from all the tables in my database?](#how-do-i-also-export-data-from-all-the-tables-in-my-database)

On this page you will find answers and guidance relating to common questions that come up when using this add-in.

## Is there a way to use a ribbon with this add-in?
Yes! Version 4.x includes a ribbon toolbar created through a twinBASIC COM add-in. The source code for the ribbon, as well as the XML definition file is included in this project. The ribbon and related files are currently on the `dev` branch until the official release of version 4.

## Why are some issues/ideas considered out of scope for this project?
This is described in more detail on [Project-Scope](Project-Scope).

## Why am I seeing a large number of "changed" files after building my project from source?
Before going into some technical details, let me clarify that in normal operation, this add-in is designed to be able to build a project with minimal, if any, changes showing between builds.

There are several possible reasons for files showing as changed even when you didn't intentionally change the source objects. Click the heading to view additional information that may be relevant in your case.

<details>
<summary><b>Form source files are showing changes in color values</b></summary>

This issue usually comes up in relation to a project being built on different computers, due to how Access internally stores the color values.

The number you see in the exported source file is affected by the current color profile and settings used by your monitor to represent the colors you see on your screen.
  
  Example:
  
  ```diff
  -     BackColor =11830108
  +     BackColor =12874308
  ```
  
</details>

<details>
<summary><b>Changes in form dimension values</b></summary>

This often happens when exporting/building on computers with different screen resolutions or monitor arrangements. These changes can often be ignored, since those values are dynamically generated. 

In most cases it would be a bit too complex to try to build the logic to determine this from the source file content, to the extend that we could discard unneeded values. One place that we have successfully done this is on the `Right` and `Bottom` dimensions of reports. (See the `SanitizeFile` function for details.)
</details>

<details>
<summary><b>Query source is significantly different</b></summary>

You may observe that the source file for a query seems to be updated to an entirely different file structure. This has to do with whether the query was saved in a compiled state in the database. If you have issues with this frequently causing changes in source files, you may want to review your workflow for editing queries. (Saving via the designer will save one way, while using the SQL view will save another way.)
</details>

<details>
<summary><b>Case changes in VBA code</b></summary>

If you see a lot of changes happening with the capitalization of variables, keyworks, properties and methods, this may be caused by the VBA editor trying to enforce consistency in the naming. This is an internal feature to VBA that some people hate and some people love. There isn't much that you can do about this behavior in the VBA IDE, but the following tips have been helpful to me in minimizing the negative effects:
* Use Pascal casing for procedures, methods and properties
* Use Hungarian notation (or similar) for variable names (i.e `lngTotal`, `strCaption`)

While many modern languages and IDE editors tend towards `camelCase` names, this just doesn't work as nicely in VBA. I personally find better success sticking with the original naming conventions the IDE was designed to work with.
  
Example ("**c**" > "**C**"):
  
  ```diff
  -    cancel = True
  +    Cancel = True
  ```
 
</details>

## How do I also export data from all the tables in my database?
Perhaps a more important question is to ask _**why**_ you want to do this... It's not that you can't, it just usually indicates a misunderstanding on the purpose of this tool. This add-in is designed to work in connection with a version control system like git to save snapshots of the _design structure_ (not the data) of a database application project. This allows you to track changes and build a copy from any point in the project's development history.

Sometimes there are pieces of data that should be included in a build, such as a table that stores configuration settings. But it can actually be very risky to commit a project's data records to version control, as this could inadvertently expose customer information (PII) or other sensitive data. It is because of this risk, and the huge problems this could cause for a well-meaning developer that we intentionally did not include a "select-all" option to easily export data from all tables.

If your underlying need is for a data backup, you might be better off reaching for another tool. (Or you could create a simple VBA script or Macro to automate the process.) If you are one of those rare but legitimate cases where you actually do need to include all the table data in version control, you will just need to go through the list of tables in the Options dialog and select each one and specify the export format. (This selection is saved with the options, and will be used on each subsequent export.) If you have a huge number of tables, you can manually edit the `options.json` file. Just be sure to follow the same format for the new entries.