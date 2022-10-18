# Excel file name and sheet name
# Be sure to close the Excel file before running the script
# Watch for background popups from Excel if errors occur. Minimize the window to see them.
[string]$ExcelFile= 'C:\Users\jasonh\OneDrive - Microsoft\Desktop\Oct 2022.xlsx'
[string]$sheetNames = @("Export")

# ADO parameters:
[string]$ADOUrl= "https://msft-skilling.visualstudio.com/"
[string]$projectName= "Content"
[string]$itemType="User Story"
[string]$areapath='Content\Production\Data and AI\Big Data, Commerce\'
[string]$iterationpath="Content"
[string]$assignee='Jason Howell ☘️'
[string]$parentItem="2919" #the ADO parent feature to link the new items to. Empty string if there is none.

## Locate the Visual Studio install on your PC
[string]$dllpath = "C:\Program Files (x86)\Microsoft Visual Studio\2017\Enterprise\Common7\IDE\CommonExtensions\Microsoft\TeamFoundation\Team Explorer"
# File 1: Microsoft.TeamFoundation.WorkItemTracking.Client.dll
# File 2: Microsoft.TeamFoundation.Client.dll

## Set mode to help set the fields that are saved into the workitems
[string]$mode="engagement"  # or "freshness"

# ADO values for Content Engagement:
if ($mode -ceq "engagement") #leave this line alone
{
[array]$tags = @('content-engagement','Scripted')
[string]$defaultDescription = "This auto-generated item was created to improve content engagement. Review <a href='https://review.learn.microsoft.com/en-us/help/contribute/troubleshoot-underperforming-articles?branch=main'>Troubleshoot lower-engaging articles</a> for tips. <br/><br/>The learn URL to improve is: "
[string]$defaultTitle= "Improve engagement: "
}

#ADO Values for Freshness:
if ($mode -ceq "freshness")  #leave this line alone
{
[array]$tags = @('content-health','Scripted')
[string]$defaultDescription = "This auto-generated item was created to track a Freshness review. Review <a href='https://review.learn.microsoft.com/en-us/help/contribute/freshness?branch=main'>the freshness contributor guide page</a> for tips. <br/><br/>The learn URL to freshen up is: "
[string]$defaultTitle= "Freshness:  "
}

#### ONLY CODE AFTER HERE
Write-Host "Script started." -ForegroundColor Cyan

# Login to Azure 
if (Get-Module -ListAvailable -Name "Az.Accounts") {
    Write-Host "Azure Module is installed. Signing in." -ForegroundColor Cyan
    $context = Get-AzContext 
    if (!$context)  { 
        Write-Host "Please sign in following the popup dialogue..."
        Connect-AzAccount 
    } 
    else  { Write-Host "   Already signed in as $($context.Account.Id). Run Disconnect-AzAccount if you need to log out." -ForegroundColor Green  }
} 
else {
    Write-Host "Sorry! Azure PowerShell needs to be installed before proceeding. Run this command:" -ForegroundColor Red
    Write-Host "Install-Module -Name Az -Scope CurrentUser -Repository PSGallery" 
    Exit
}


try{
    $Excel = New-Object -ComObject Excel.Application 

    Write-Host "Trying to open Excel file. If the script hangs, minimize PowerShell and check for popup messages in Excel." -ForegroundColor Cyan
    # Open the Excel file
    $Workbook = $Excel.Workbooks.Open($ExcelFile)
   
    # We'll store the Excel data in hash tables
    [hashtable]$columnHashtable= [ordered]@{};
    [hashtable]$rowHashtable = [ordered]@{};
    [array]$allRowsArray= @();


    ## Loop over multiple worksheets in the list if needed
    foreach($sheetName in $sheetNames)
    {
        $sheet = $workbook.Sheets.Item($sheetName)
        Write-Host "Opened Excel workbook named:" $sheetName
    
        $sheet.Columns.AutoFit();

        Write-Host "Started parsing header columns." -ForegroundColor Cyan
        ## Header parsing loop to discover column header names
        $rownumber = 1
        $columnNumber = 1
    
        # Loop over each column until there is an empty heading
        do {
            $columnHashtable.Add($columnNumber, $sheet.Rows[$rownumber].Cells[$columnNumber].Text)
            $columnNumber++
        }
        while ($sheet.Rows[$rownumber].Cells[$columnNumber].Text -ne "")
    
        # display the columns for debugging if needed
        $columnHashtable.GetEnumerator() | Sort-Object Name
    
        Write-Host "Finished parsing header columns." -ForegroundColor Cyan
    
        ## Parse the data rows from top down until an empty row is reached
        Write-Host "Starting parsing data rows."  -ForegroundColor Cyan

        $rownumber = 2 #Starting row to parse data. 
    
            do {
    
                #clear the row in each loop iteration
                $rowHashtable = [ordered]@{}
                Write-Host "Started parsing row: $rownumber for URL: " $sheet.Rows[$rownumber].Cells[4].Text

                # Loop over each column in the row and record the text value into the hashtable
                for($columnNumber = 1; $columnNumber -lt $columnHashtable.Count+1; $columnNumber++)
                {
                   # for debugging if needed: write-host $columnNumber
                   $rowHashtable.Add( $columnHashtable[$columnNumber], $sheet.Rows[$rownumber].Cells[$columnNumber].Text)
                   #reading cell color is possible but not yet implemented
                   #$sheet.Rows[$rownumber].Cells[$columnNumber].Interior.ColorIndex
                   #https://learn.microsoft.com/en-us/office/vba/api/Excel.ColorIndex

                   # for debugging if needed: $rowHashtable

                }
        
                # Save each row into an array
                $allRowsArray += $rowHashtable
                
                Write-Host "Finished parsing row: $rownumber for URL: " $sheet.Rows[$rownumber].Cells[4].Text
        
                #Looping to the next row       
                $rownumber++

            }

            # exit the loop when an empty cell/row is found
            while ($sheet.Rows[$rownumber].Cells[1].Text -ne "")
            Write-Host "Finished parsing data rows."  -ForegroundColor Cyan
    
        }
}
catch 
{ 
    Write-Host "Couldn't parse the Excel file" 
    Write-Host $_
    Write-Host $_.ScriptStackTrace
    Exit
}
finally 
{
    $Excel.Workbooks.Close()
}

try{

    Write-Host "Loading Devops dlls."  -ForegroundColor Cyan
    # load the required dll
    Add-Type -path "$dllpath\Microsoft.TeamFoundation.WorkItemTracking.Client.dll"
    Add-Type -path "$dllpath\Microsoft.TeamFoundation.Client.dll"

    # connect to the ADO instance
    Write-Host "Connecting to Azure DevOps..." -ForegroundColor Cyan
    $vsts = [Microsoft.TeamFoundation.Client.TfsTeamProjectCollectionFactory]::GetTeamProjectCollection($ADOUrl)
    $WIStore=$vsts.GetService([Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemStore])
    $project=$WIStore.Projects[$projectName]
    $type=$project.WorkItemTypes[$itemType]
    
    # Loop over all rows, and create a workitem for each row.
    foreach($row in $allRowsArray)
    {
        Write-Host "Processing row " $row.url -ForegroundColor Cyan
  
        Write-Host "Creating a Azure DevOps Workitem as follows:" -ForegroundColor Cyan

        # customize which fields you want in the short table
        [string]$description = ""
        $description += $defaultDescription
        $description += "<br/><a href={0} target=_new>{0}</a><br/>" -f $row["Url"] 
        $description += "<table style='border: 1px solid black; border-collapse: collapse;'>"

        if ($mode -ceq "freshness")
        {
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Freshness</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["Freshness"]
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>LastReviewed</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["LastReviewed"]
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>MSAuthor</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["MSAuthor"]
        }
        if ($mode -ceq "engagement")
        {
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Engagement</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["Engagement"]
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Flags</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["Flags"]
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>BounceRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["BounceRate"]
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>ClickThroughRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["ClickThroughRate"]
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>CopyTryScrollRate</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["CopyTryScrollRate"] 
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>Freshness</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["Freshness"]
            $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>LastReviewed</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row["LastReviewed"]
        }

        $description += "</table><br/>"

        $description += "Other page properties:<br/>"
        $description += "<table style='border: 1px solid black; border-collapse: collapse;'>"

        # Loop over all fields in the row to create the additional detailed table. Sorted alphabetically for now.
        foreach($keyname in ($row.keys | Sort-Object ))
            {
                # handle URLs as links
                if ($keyname -ceq "Drilldown" -or $keyname -ceq "Trends" -or $keyname -ceq "GitHubOpenIssuesLink")
                { 
                    $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>$keyname</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'>"
                    if (-not [string]::IsNullOrWhiteSpace($row[$keyname]))
                    { 
                        $description += "<a href='"
                        $description += $row[$keyname]
                        $description += "' target='_new'>URL</a></td></tr>"
                    }
                    else 
                    {
                        $description += "</td></tr>"
                    }
                }
                # handle others fields as text
                else
                {
                    $description += "<tr><td align='right' style='border: 1px solid black; border-collapse: collapse;'><strong>$keyname</strong></td><td align='left' style='border: 1px solid black; border-collapse: collapse;'> {0}</td></tr>" -f $row[$keyname]
                }
            }
        $description += "</table>"

        # Set new workitem properties
        $item = new-object Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem $type
        $item.Title = $defaultTitle + $row["Url"]
        $item.AreaPath = $areapath
        $item.IterationPath = $iterationpath
        $item.Tags = ($tags |Select-Object) -join ","
        $item.Description = $description
        $item.Fields['Assigned To'].Value = $assignee
        $item.save()

        # Link to a parent item if there's a number to link to
        if (-not [string]::IsNullOrWhiteSpace($parentItem))
        {
            Write-Host "Adding link to workitem parent"
            $hierarchicalLink = $wiStore.WorkItemLinkTypes["System.LinkTypes.Hierarchy"];
            $workitemlink = new-object Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItemLink $hierarchicalLink.ReverseEnd, $($parentItem)
            $item.WorkItemLinks.Add($workitemlink)
            $item.save()
        }

        # show the workitem properties as output for confirmation
        $item | Select-Object Id,AreaPath,IterationPath,@{n='AssignedTo';e={$_.Fields['Assigned To'].Value}},Title, Tags
               
        Write-Host "Saved the workitem to URL: $($ADOUrl)$($projectName)/_workitems/edit/$($item.Id)" -ForegroundColor Green
        Write-Host ""
    }

    # Open the recent Devops items in a browser
    Start-Process "$($ADOUrl)$($projectName)/_workitems/recentlyupdated/"

}
catch 
{ 
    Write-Host "Couldn't create the AzureDevOps items." 
    Write-Host $_
    Write-Host $_.ScriptStackTrace
}
finally
{
    Write-Host "Done!" -ForegroundColor Green

}
    