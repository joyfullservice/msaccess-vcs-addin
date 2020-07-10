We've got issues
---------
The easiest way to contribute is to make a detailed Issue. Be sure to include details about the version of OS, Access, and VCS add-in. 
If you can, provide a [Minimal Reproducible Example](https://stackoverflow.com/help/minimal-reproducible-example) of the problem.

Development Version
---------
If you want the very latest working copy you will have to make from GitHub to git to Access:

* Download the latest release.
* Install the add-in. (Just open *Version Control.accda*)
* Clone this repository.
* Pull your clone down to your local machine.
* Choose a branch in git. (Either master, or the most recent development branch)
* Use the add-in to build from source. 
* Run the newly compiled *Version Control.accda* file to install the development version

Making your first Pull Request (PR)
---------
Ok, so you want to change the code and submit your changes back to the community. How kind of you! If you followed the steps to get to the Development Version then you can follow these steps to go back from Access to git back up to GitHub:
* **Always create a new branch when you want to make a change.** Make the name of the branch descriptive to what you set out to accomplish. 
* Open the development copy of *Version Control.accda* 
* Perform testing on your development version.
* Make updates to the project. (This is where the magic happens)
* When you are ready to make a commit run the **Deploy** procedure by typing `Deploy` in the the VBA immediate window and pressing **Enter**. This will:
   * Increment the version number.
   * Export the project to source.
   * Install the version you have open. 
* Close Access.
* Open an Access project and ensure that the version installed matches the new version you just deployed. 
* Perform testing to confirm that your new version works as expected. 
* Make a git commit and briefly describe your changes in the commit notes. (You can add more verbose details in your pull request.)
* Push your branch up to your cloned repository.
* Make a pull request to the upstream project! Be sure to clearly describe what you did and why in the pull request. This will allow reviewers to better understand why your PR should be merged.
