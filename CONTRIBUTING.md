Development Version
---------
If you want the very latest working copy you will have to:
    * Download the latest release
    * Install the add-in
    * Clone this repository
    * Pull your clone down to your local machine
    * Choose a branch in git (Either master, or the most recent development branch)
    * Use the add-in to build from source 
    * Run the newly compiled .accda file and it will install the development version

Making your first Pull Request (PR)
---------
Always create a new branch when you want to make a change. Make the name of the branch descriptive to what you set out to accomplish. 
Open the development copy of *Version Control.accda* and perform testing on your development version. 
Make updates to the project. 
When you are ready to make a commit run the **Deploy** procedure by simply typing Deploy in the the VBA immediate window and pressing **Enter**. This will increment the version number, Export the project to source, and install the version you have open for your profile. Close Access.
Open an Access project and ensure that the version installed matches the new version you just deployed. Perform testing to confirm that your new version works as expected. 
Make a git commit and briefly describe your changes in the commit notes. You can add more verbose details in your pull request.
Now you can push your branch up to your cloned repository and make a pull request to the upstream project! Be sure to clearly describe what you did and why in the pull request. This will allow reviewers to better understand why your PR should be merged.