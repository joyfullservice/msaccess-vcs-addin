So, you like the Addin. And you want to contribute more. Hurrah!

# BLUF:
1. Make Pull Requests (PRs) on the "Dev" branch: it's the in use branch. Stable is not actively worked on, and is used to provide a "stable" base while we work out any kinks in the dev branch.
2. If you can, do not pull directly from your "dev" fork: make a working branch in your repository. This will help ensure reduced conflicts, and to ensure we know what the scope of your PR is more easily. 
3. Keep scope of PRs within a single area of focus: if you fixed two bugs, please keep them to separated PRs. Using specific work branches will help. This will ensure we don't get co-mingled issues, and is a lot cleaner to ensure we don't introduce new bugs from fixing others.

# Details:
If you directly edit Access Add-ins (such as this one) within the "opening" Access file, changes will not be saved. 
This is a double edged sword: it allows easy debugging, and trying things out which might otherwise ruin files. Downside is that once you close the session, it will discard any settings. 

This is a nice  way to load "extras" for users and ensure they don't break things for everyone else. 

If you want to make changes to your Addin (and contribute those to others!), do this:
1. Fork this [MS Access Repository](https://github.com/joyfullservice/msaccess-vcs-integration) into your GitHub account repos.
![image](https://user-images.githubusercontent.com/54177882/117137254-6d378280-ad77-11eb-923e-a7a876611fed.png)
2. Clone your fork to a local repository alongside your other Access dev repos on your machine. 
3. Some put theirs alongside some other Access repositories they utilize.

![image](https://user-images.githubusercontent.com/54177882/117137620-f353c900-ad77-11eb-9680-047cabd002da.png)

3. Switch to "Dev" Branch: We suggest renaming YOUR "dev" branch to something local to you, especially if you still use some of the legacyVCS tools. I named my local fork of this branch to "dev-addin".
4. Connect a second remote to the joyfullservice/msaccess-vcs-integration (upstream) repository. This way you can track progress on the upstream ( joyfullservice/msaccess-vcs-integration) repository. There may be changes you don't want to pull into yours, or you may want customization not present on the upstream (in some environments, users have specific changes required to ensure proper integration in their security environment and/or configurations that shouldn't be default for everyone).

![image](https://user-images.githubusercontent.com/54177882/117138802-84776f80-ad79-11eb-97f0-e55e62f59c38.png)


5. Go into your new local repo, and launch the Addin directly. 
6. Make changes, use the add-in to export the add-in code, and commit/push/pull request just like any other repository.

![image](https://user-images.githubusercontent.com/54177882/117139316-197a6880-ad7a-11eb-95ca-1cb3c12a712f.png)

7. To edit the VCS, click the "x" button instead of the "install" on the loading form.

![image](https://user-images.githubusercontent.com/54177882/117144981-990b3600-ad80-11eb-8413-db75258dc9ca.png)

8. To edit forms for the VCS, open in "Design View"; their loading code / functions won't run.

![image](https://user-images.githubusercontent.com/54177882/117144997-9c9ebd00-ad80-11eb-8c31-a56ed881fc18.png)


9. See above for PR guidelines.
