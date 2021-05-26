- [Is there a way to use a ribbon with this add-in?](#is-there-a-way-to-use-a-ribbon-with-this-add-in)
- [Why are some issues/ideas considered out of scope for this project?](#why-are-some-issuesideas-considered-out-of-scope-for-this-project)
- [Why am I seeing a large number of "changed" files after building my project from source?](#why-am-i-seeing-a-large-number-of-changed-files-after-building-my-project-from-source)
- [Will this fork eventually be merged back into the upstream repository?](#will-this-fork-eventually-be-merged-back-into-the-upstream-repository)

On this page you will find answers and guidance relating to common questions that come up when using this add-in.

## Is there a way to use a ribbon with this add-in?
We would love to use a custom ribbon to make the controls and features of this add-in more intuitive, but thus far I have been unable to create a working com add-in for Microsoft Access. See [issue #34](https://github.com/joyfullservice/msaccess-vcs-integration/issues/34) for more details.

## Why are some issues/ideas considered out of scope for this project?
This is described in more detail on [[Project-Scope]].

## Why am I seeing a large number of "changed" files after building my project from source?
Before going into some technical details, let me clarify that in normal operation, this add-in is designed to be able to build a project with minimal, if any, changes showing between builds.

There are several possible reasons for files showing as changed even when you didn't intentionally change the source objects. Click the heading to view additional information that may be relevant in your case.

<details>
<summary><b>Form source files are showing changes in color values</b></summary>

This issue usually comes up in relation to a project being built on different computers, due to how Access internally stores the color values. 

The number you see in the exported source file is affected by the current color profile and settings used by your monitor to represent the colors you see on your screen.
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
</details>


## Will this fork eventually be merged back into the upstream repository? 
The **joyfullservice** branch is a near complete rewrite of the original project. It is unlikely that it will ever be merged back into the upstream project of `msaccess-vcs-integration/msaccess-vcs-integration`. This upstream link is primarily maintained to give visibility to this branch for those that may be searching for an add-in based version control solution for Microsoft Access.
