# Using Version Control System

In the process of developing a more complex Microsoft Access database, you may find yourself wondering what changed when, or how you are going to collaborate with other developers. That's what this system was designed for. It basically exports your Access database into individual files representing the various tables, queries, forms, etc... in such a way that you can easily compare changes or revisions in the development process. Let's consider a couple use cases:

## Export and Save
For some developers, they simply want to **track the progress** of their work on a database system over time. Using this Version Control System, you can simply **export** the project to source files, then commit those changes to GitHub (or other versioning system). If you are looking for something simple, GitHub Desktop provides a nice user interface for working with a GitHub/GitLab repository.

## Collaborative Development
In other cases you might have a team of developers that are simultaneously working on the same Microsoft Access Database system. This Version Control System allows developers to work independently on their own copies of the database, **exporting** and **committing** changes to a repository. Changes are reviewed and **merged at a source file level**, and then the database is **built** from the source files to combine all the changes together.

_Note: This aspect of building the database from source files is still under development._

# Options

The options dialog can be opened from the main screen by click the `Options` button. Internally the options are stored in a `vcs-options.json` file in the path of the exported source code. These options are loaded and used when exporting to source files, or when building a project from source.

[Click here for detailed Options Information](Options).
