# Signature Add-in

This is a modified Signature Add-in based from the one provided by the [Microsoft team here](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature). The way this has been modified:

- Javascript ES6 (classes)
- Updated Webpack to export based on classes
- Reduction in redundant code
- Reduction in HTML pages needed
- Consolidation of CSS
- Comments and JSDoc Comments

## Configure VSCode and Run Add-in

To configure this project for VSCode:

1. Setup VSCode:

- Install [VSCode](https://code.visualstudio.com/).
- Install [Node.js](https://nodejs.org/en/download).

1. Download the code as a ZIP.
1. Extract to a folder on your computer.
1. Open VSCode, **File**, **Open Folder** and open the extracted folder.
1. Press CTRL+~ to open the Terminal window.
1. Type: npm install.
1. Once it is complete, type: npm start.
1. This will open a Terminal window, open Outlook.com in your browser.
1. Select an email, click the drop-down menu **(...)**, select **Get Add-ins**.
1. Click **My Add-ins**.
1. In the *Custom Add-ins section*, click **Add a custom add-in**, and select **Add from File...**
1. Browse to an select the "manifest.xml" file in the extracted folder and click Install on the dialog that pops up.
1. Open a new email, you will see a prompt at the top: Please set your signature with the Office Add-ins sample. Set signatures | Dismiss
1. Click **Set Signatures**

## Questions

Please send questions my way [via by blog](https://theofficecontext.com).
