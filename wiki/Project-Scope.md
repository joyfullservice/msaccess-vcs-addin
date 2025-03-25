I personally believe that one of the keys to the long-term success of this project is establishing a clearly defined scope of what it is intended to do, and sticking to that scope.

## Guiding Principles
Since much of this could be considered subjective archetecutual design decisions, let me outline some of the guiding principles that I am trying to keep in the forefront of this project.
* The **fundemental purpose** of this add-in is two-fold:
  * **Export** database objects as source files
  * **Import**/**Build**/**Merge** source files to database objects
* The goal is to replicate the original database as **closely as possible** when building from source.
* The add-in is **not intended** to fix/repair/enhance the target database, other than what is necessary to perform the basic functions of exporting and importing source files.
* The **user interface** should be as intuitive and user-friendly as possible. It should be both efficient for the expert, and easy for the beginner. Flexibility without clutter.
* The tool should be **extensible**, where internal code can be added to carry out additional tasks outside the scope of this add-in.


## How Features are Evaluated
Features add complexity, and complexity increases [cost of carry](https://martinfowler.com/bliki/Yagni.html). Great features are welcomed and make this tool better every year. Unecessary or overly complex features bog down the project and slow progress on more important issues.

**Feature considerations:**
* How many users does this affect? Will this benefit everyone, or just a single user with a really unique setup?
* How complex is the feature? Is it limited to changes in a few areas of code, or are we talking about a significant refactoring?
* Do functionality changes cause any risks for those currently using the add-in in production environments?

If your idea didn't get implemented, don't take it personally.  :-)  Remember, this is an ongoing work in progress, and someone has to make the hard decisions about what gets added and what doesn't.
