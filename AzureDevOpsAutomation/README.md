# Create Azure DevOps workitems from a Excel worksheet

This PowerShell script  automates entering maintenance items into our Azure DevOps boards.
![Screenshot of the script running in Visual Studio Code](https://user-images.githubusercontent.com/5067358/195956233-feac7ab6-0a9f-437c-8473-8fa2752c5df1.png)

## Gather maintenance data about content assets

1. Open the Content Engagement Report [http://aka.ms/contentengagementreport](http://aka.ms/contentengagementreport)

2. On the Documentation tab, set the filters to focus on the docs you want to work on.

3. **Export data** from Content Engagement report to download an .xlsx file.

4. Note the path of the file to use in the script later, and the workbook (tab) name.

5. Unblock the downloaded file: From Windows Explorer, right-click on the file, choose **properties**, select **unblock**, and **apply**).

6. Edit the Excel worksheet to keep the rows you want to make work items for, removing any unwanted data.

   - Delete any extra rows you don't want to create items for. For example, you may want to prioritize 1000+ page views data and remove the rest.
   - Leave the column headers in tact, since the script will parse those dynamically.
   - Adjust the width of any narrow columns in Excel that are hiding the data with `######` and leave a buffer of extra white space on those columns for best results.

7. Save and close the Excel file.

## Run the PowerShell script

1. Download the script and make sure the file name has the .ps1 extension.

2. Launch **Windows PowerShell ISE** or **Visual Studio Code** IDE from the Start menu. Open the script file.

   You may need to install the Code [Extension for PowerShell](https://marketplace.visualstudio.com/items?itemName=ms-vscode.PowerShell) to run it interactively

3. From the terminal, run the following command so that the script isn't blocked from executing.

   `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`

4. Edit the string parameters at the top of the script to customize to your liking. Set the mode to do freshness or content engagement settings.

5. Play the script to parse the Excel file and create work items automatically. Review the Output window to see the progress or any error messages that appear.

6. After it runs, a browser window should launch to help you find the newly created work items, or you can review the output window for a list of the created items and their ID numbers.

## Prerequisites

1. The script uses a .dll from Visual Studio 2017 or later interface with DevOps. IF you have another version of Visual Studio installed, it may work OK. Just update the DLL path in the script.

   `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer`
   - Microsoft.TeamFoundation.WorkItemTracking.Client.dll
   - Microsoft.TeamFoundation.Client.dll

2. The script prompts you to log in with your Azure credentials.
Note:  Creating tags in the ADO instance requires additional permissions that may cause a failure if you add new tags that are not yet in the system.

## Wishlist to-do

1. Add color coding to the low engagement fields. Its hard to do since the current Excel export doesn't include colors or an indicator of what high and low engagement is.

2. Detect duplicate items that already exist in DevOps and skip creating those.

3. Assign the new maintenance items to article ms.author value (optional)

## Known issues

1. If you don't have Azure PowerShell installed, you can install with this command or run the MSI

   `Install-Module -Name Az -Scope CurrentUser -Repository PSGallery`

2. If there is an authentication failure you'll get this error. You need to connect to Azure successfully using cmdlet Get-AzContext.

   ```output
   Exception calling "GetService" with "1" argument(s): "TF30063: You are not authorized to access https://msft-skilling.visualstudio.com/."  at <ScriptBlock>, C:\Downloads\CreateWorkitemsFromExcelFile.ps1: line 128
   ```

3. Investigating this error

   ```output
   Couldn't create the AzureDevOps items.
   Exception calling "GetTeamProjectCollection" with "1" argument(s): "Could not load type 'System.Diagnostics.Eventing.EventProvider' from assembly 'System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'."
   ```

4. Data appears as `#####` symbols in the ADO item when the source column in Excel was too narrow to parse. The script attempts to resize the columns to make it fit. You can expand the column width to include extra whitespace, save the Excel file, close the file, and try the import again.

   ![image](https://user-images.githubusercontent.com/5067358/195960464-7f4bb326-a5ea-43fa-9b92-b4c76788f54f.png)
