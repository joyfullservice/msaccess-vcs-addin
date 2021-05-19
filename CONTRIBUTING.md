Found an Issue? Have an idea?
---------
The easiest way to contribute is to create a detailed [Issue](https://github.com/joyfullservice/msaccess-vcs-integration/issues). Be sure to include details about the version of OS, Access, and VCS add-in. 
If you can, provide a [Minimal Reproducible Example](https://stackoverflow.com/help/minimal-reproducible-example) of the problem.

Development Version
---------
If you want the very latest updates since the last published release, you will need to build it from source. Here is how to go from GitHub to git, to Access:

* Make sure you have a fairly recent version of the add-in installed. If not:
  * Download the latest release.
  * Install the add-in. (Just open *Version Control.accda*.)
* Clone this repository.
* Pull your clone down to your local machine.
* Choose a branch in git. (Typically `dev`)
* Use the add-in to *Build From Source*, selecting the cloned `Version Control.accda.src` folder.
* Run the newly compiled *Version Control.accda* file to install the development version.

Making your first Pull Request (PR)
---------
A *Pull Request* is how you can propose that your code changes be included in the main project. (This project is the work of many people who have donated their efforts to make it better for everyone.) If you followed the steps to get to the Development Version then you can follow these steps to go back from Access to git, back up to GitHub:
* *Optional: For larger changes, you should consider making a branch that describes the changes you are proposing.*
* Open the development copy of *Version Control.accda* from the cloned GitHub project.
* Perform testing on your development version.
* Make updates to the database project. (This is where the magic happens)
* When you are ready to make a commit run the **Deploy** procedure by typing `Deploy` into the VBA immediate window and press **Enter**. This will:
   * Increment the version number.
   * Export the project to source.
   * Install the version you have open. 
* Close Access.
* Open an Access project and ensure that the version installed matches the new version you just deployed.
* Perform testing to confirm that your new version works as expected. 
* Make a git **commit*** and briefly describe your changes in the commit notes. (You can add more verbose details in your pull request.) 
  * _*When creating the commit, please select only the files that reflect the actual changes you are proposing. It is usually not necessary to include auto-generated files that don't include substantive or intended changes._
* **Push** your branch up to your cloned repository.
* Make a **pull request** to the upstream project! Be sure to clearly describe what you did and why in the pull request. This will allow reviewers to better understand why your PR should be merged.
* Make PR on the `dev` branch: it's the most active. `Master` branch PRs should be limited to critical bug fixes (and should also be addressed in `dev`, too). 
* *Tip: If you have many different types of changes to propose, please use different pull requests for each of them. That will be easier to review and implement them individually.*

Thank you again for your support for the Microsoft Access development community!!
