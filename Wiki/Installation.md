# Install

1. Download the [latest **Version_Control_vX.X.X.zip**](https://github.com/joyfullservice/msaccess-vcs-addin/releases/latest).
2. Extract `Version Control.accda` from the zip archive.
3. Remove "Mark of the Web" (MOTW) if Windows blocks the file. See [this blog post](https://nolongerset.com/install-vcs-addin/). MOTW is not always set (for example when extracting with 7-Zip).
4. *(Optional)* Adjust install options (below).
5. Click **Install Add-In**.

![Install Form Image](img/install.png)

## Install Options

None of these are required for basic use; they help locked-down environments.

| Setting | Default | Description |
|---------|---------|-------------|
| **Trust Add-In Folder** | On | Adds the install folder as a trusted location (`%AppData%\Roaming\MSAccessVCS\` by default). |
| **Advanced Options** | Off | Ribbon, trust-on-open, custom install path. |

![Install Form Image](img/install-advanced.png)

| Setting | Default | Description |
|---------|---------|-------------|
| **Use Ribbon Addin** | On | Installs the twinBASIC COM ribbon add-in (32- and 64-bit DLLs). Provides **Export**, **Build**, **Merge**, **Options**, **Run Tests**, and related commands. Turn off only if COM add-ins are blocked; use **Database Tools** → **Add-Ins** → **VCS** menu items instead. |
| **Open add-in after installing to trust the add-in file** | Off | Opens the `.accda` after install so you can trust it in strict environments. |
| **Create compiled version** | Off | Installs a compiled `.accde` instead of the `.accda`. |
| **Use helper script (Worker.vbs)** | On | Allows the add-in to run a small helper script for the few jobs it cannot do inside its own process. Turn it off if antivirus or endpoint protection blocks it (below). |
| **Install Folder** | `%AppData%\Roaming\MSAccessVCS\` | Custom path if policy requires it. **Uninstall completely before changing** the install folder. |

## Worker script

A few jobs cannot be done from inside the process that is running the add-in: deleting the add-in file while Access still has it open, checking whether other users can open your database, saving the VBA project, and rebuilding the add-in itself. For these the add-in writes a small script (`Worker.vbs`) into the install folder and runs it.

Some endpoint protection products treat Access launching a freshly written script as suspicious and block it, which can leave the add-in unable to finish these steps. If that describes your environment, uncheck **Use helper script (Worker.vbs)** in **Advanced Options** during install. The script is deleted from the install folder when you do.

Everything else keeps working. What changes:

| Job | With the script off |
|-----|---------------------|
| **Export / merge** | If your VBA project has unsaved changes, the add-in tells you to press **Save** in the Visual Basic Editor and run the export again, instead of saving it for you. Exported source would otherwise be missing your latest form and report code. |
| **Build and merge** | The database is always closed and reopened afterwards, because there is no way to check whether it is still available to other users. Slightly slower; no difference in the result. |
| **Uninstall** | Access closes and lists the files for you to delete by hand, instead of removing them itself. |
| **Rebuild Add-In** | Unavailable. Build from source and open the resulting file to install it. |

To turn it back on, re-run the installer with **Use helper script (Worker.vbs)** checked.

## Ribbon (Version 4+)

The ribbon is a lightweight COM add-in (`MSAccessVCSLib_win32.dll` / `MSAccessVCSLib_win64.dll`) that calls into `Version Control.accda`. It loads when Access starts if installation succeeded and COM add-ins are enabled in Trust Center.

If the ribbon does not appear:

- Confirm **Use Ribbon Addin** was on during install.
- Check **File** → **Options** → **Trust Center** → **COM Add-ins**.
- Re-run the installer or use **Options** → **Settings** → **Open Install Folder** to verify DLLs are present.

## Rebuild add-in from source (developers)

Contributors building the add-in from `Version Control.accda.src` can use **Rebuild Add-In** on the ribbon (or `VCS.RebuildAddIn` in code) after a local build. See [Editing and Contributing](Editing-and-Contributing).

## Uninstall

[Uninstall instructions](Options#remove-add-in) — **Options** → **Settings** → **Uninstall**, or the wiki Options page under **Remove Add-In**.
