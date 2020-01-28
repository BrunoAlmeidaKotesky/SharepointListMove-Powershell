#Function to copy attachments between list items
Function Copy-SPOAttachments($SourceItem, $TargetItem)
{
    Try {
        #Get All Attachments from Source
        $Attachments = Get-PnPProperty -ClientObject $SourceItem -Property "AttachmentFiles"
        $Attachments | ForEach-Object {
        #Download the Attachment to Temp
        $File  = Get-PnPFile -Url $_.ServerRelativeUrl -FileName $_.FileName -Path $env:TEMP -AsFile -force
 
        #Add Attachment to Target List Item
        $FileStream = New-Object IO.FileStream(($env:TEMP+"\"+$_.FileName),[System.IO.FileMode]::Open) 
        $AttachmentInfo = New-Object -TypeName Microsoft.SharePoint.Client.AttachmentCreationInformation
        $AttachmentInfo.FileName = $_.FileName
        $AttachmentInfo.ContentStream = $FileStream
        $AttachFile = $TargetItem.AttachmentFiles.add($AttachmentInfo)
        $Context.ExecuteQuery()   
     
        #Delete the Temporary File
        Remove-Item -Path $env:TEMP\$($_.FileName) -Force
        }
    }
    Catch {
        write-host -f Red "Error Copying Attachments:" $_.Exception.Message
    }
}
 
#Function to list items from one list to another
Function Copy-SPOListItems()
{
    param
    (
        [Parameter(Mandatory=$true)] [string] $SourceListName,
        [Parameter(Mandatory=$true)] [string] $TargetListName
    )
    
        #Get All Items from the Source List in batches
        Write-Progress -Activity "Reading Source..." -Status "Getting Items from Source List. Please wait..."
        $SourceListItems = Get-PnPListItem -List /Lists/ListaLivroGP
        $SourceListItemsCount= $SourceListItems.count
        Write-host "Total Number of Items Found:"$SourceListItemsCount       
 
        #Get fields to Update from the Source List - Skip Read only, hidden fields, content type and attachments
        $SourceListFields = Get-PnPField -List /Lists/ListaLivroGP;
     
        #Loop through each item in the source and Get column values, add them to target
        [int]$Counter = 1
        ForEach($SourceItem in $SourceListItems)
        { 
            $ItemValue = @{}
            #Map each field from source list to target list
            Foreach($SourceField in $SourceListFields)
            {
                #Check if the Field value is not Null
                If(!$SourceItem[$SourceField.InternalName])
                {
                    #Handle Special Fields
                    $FieldType  = $SourceField.TypeAsString
 
                   
                        #Get Source Field Value and add to Hashtable
             $ItemValue.add($SourceField.InternalName,$SourceItem[$SourceField.InternalName])
                
                }
            }
            Write-Progress -Activity "Copying List Items:" -Status "Copying Item ID '$($SourceItem.Id)' from Source List ($($Counter) of $($SourceListItemsCount))" -PercentComplete (($Counter / $SourceListItemsCount) * 100)
         
            #Copy column value from source to target
            $NewItem = Add-PnPListItem -List $TargetListName -Values $ItemValue
 
            #Copy Attachments
            Copy-SPOAttachments -SourceItem $SourceItem -TargetItem $NewItem
 
            Write-Host "Copied Item ID from Source to Target List:$($SourceItem.Id) ($($Counter) of $($SourceListItemsCount))"
            $Counter++
        }

}
 
#Connect to PnP Online
Connect-PnPOnline -Url "https://trentim.sharepoint.com" -Credentials (Get-Credential)
$Context = Get-PnPContext
 
#Call the Function to Copy List Items between Lists
Copy-SPOListItems -SourceListName "ListaLivroGP" -TargetListName "testeDavi"


#Read more: https://www.sharepointdiary.com/2017/01/sharepoint-online-copy-list-items-to-another-list-using-powershell.html#ixzz6CKlMgc5r