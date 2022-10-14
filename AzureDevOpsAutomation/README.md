# Create Azure DevOps workitems from a Excel worksheet

This PowerShell script  automates entering maintenance items into our Azure DevOps boards.

## Gather maintenance data about our articles

1. Open the Content Engagement Report [http://aka.ms/contentengagementreport](http://aka.ms/contentengagementreport)

2. On the Documentation tab, set the filters to focus on the docs you want to work on.

3. **Export data** from Content Engagement report to download an .xlsx file.

4. Note the path of the file to use in the script later, and the workbook (tab) name.

5. Unblock the downloaded file: From Windows Explorer, right-click on the file, choose **properties**, select **unblock**, and **apply**).

6. Edit the Excel worksheet to include only the rows you want to make work items for. Delete any extra rows you don't want items for. Leave the column headers. Adjust the width of any narrow columns in Excel that block the data with `######` and leave extra white space for best results.

7. Close the Excel file.

## Run the PowerShell script

1. Download the script and make sure the file name has the .ps1 extension.

2. Launch **Windows PowerShell ISE** or **Visual Studio Code** IDE from the Start menu. Open the script file.

3. Edit the string parameters at the top of the script to customize to your liking.

4. Play the script to parse the Excel file and create work items automatically. Review the Output window to see the progress or any error messages that appear.

5. After it runs, a browser window should launch to help you find the newly created work items, or you can review the output window for a list of the created items and their ID numbers.

## Prerequisites

1. The script uses a .dll from Visual Studio 2017 or later interface with DevOps. IF you have another version of Visual Studio installed, it may work OK. Just update the DLL path in the script.

   `C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer`
   - Microsoft.TeamFoundation.WorkItemTracking.Client.dll
   - Microsoft.TeamFoundation.Client.dll

2. The script prompts you to log in with your Azure credentials.
Note:  Creating tags in the ADO instance requires additional permissions that may cause a failure if you add new tags that are not yet in the system. 

## Wishlist to-do

1. Resize fields in Excel automatically to avoid dates with `######` as the value

2. Add color coding to the low engagement fields. Its hard to do since the current Excel export doesn't include colors or an indicator of what high and low engagement is.

3. Detect duplicate items that already exist in DevOps and skip creating those.

4. Assign the new maintenance items to article ms.author value (optional)

## Known issues

1. If you don't have Azure PowerShell installed, you can install with this command or run the MSI

   `Install-Module -Name Az -Scope CurrentUser -Repository PSGallery`

2. If there is an authentication failure you'll get this error. You need to connect to Azure succesfully.

   ```output
   Exception calling "GetService" with "1" argument(s): "TF30063: You are not authorized to access https://msft-skilling.visualstudio.com/."  at <ScriptBlock>, C:\Downloads\CreateWorkitemsFromExcelFile.ps1: line 128
   ```

3. Investigating this error

   ```output
   Couldn't create the AzureDevOps items.
   Exception calling "GetTeamProjectCollection" with "1" argument(s): "Could not load type 'System.Diagnostics.Eventing.EventProvider' from assembly 'System.Core, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089'."
   ```

4. Data appears as `#####` symbols in the ADO item when the source column in Excel was too narrow to parse. You can expand the column width to include extra whitespace, save the Excel file, close the file, and try the import again.
