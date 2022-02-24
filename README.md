# Outlook AddIn: Move By Category
Outlook AddIn that moves mails to folders based on assigned category.
Tested on Outlook 2013, might work on other versions.

# Install
Unzip and run setup.
Restart Outlook.

# Usage (manual)
AddIn will add ribbon tab ARCHIVE.

![Alt text](/img/ribbon.png?raw=true "Ribbon")

Click on "Expand" (lower right corner) will open config window with list of all Categories.
![Alt text](/img/config.png "Config")

Dobule click on category to select archive destination.

After configuration, select Archive from ribbon or Categorize from context menu to move selected items to archive folders depending on item category.

![Alt text](/img/context.png?raw=true "Context menu")

# Usage (from Rules)
Add script from VBA folder to Outlook.
Create rule with action "run a script" and select previous script.
When rule is triggered, mail will be moved to folder based on assigned category.
