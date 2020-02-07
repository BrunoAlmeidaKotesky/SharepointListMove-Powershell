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







function userOptions {
    param([bool]$option = $true)
    $Site = Read-Host 'Qual a url do site de onde a lista será copiada?';
    #While pra validar se url é vazia
    while (!$Site) {
        $Site = Read-Host 'Qual a url do site de onde a lista será copiada?';
    }
    if($false -eq $option){
        $Site2 = $null;
    }
    else{
        $Site2 = Read-Host 'Qual a url do site para qual a lista será criada?';
        #While pra validar se url é vazia
        while (!$Site2) {
            $Site2 = Read-Host 'Qual a url do site para qual a lista será criada?';
        }
    }
    
    $result = retryConnection -siteurl $Site -isExternal $true;
    if ($result -eq "UriFormatException" -or $result -eq "WebException" -or $result -eq "IdcrlException") {
        Write-Host "As credenciais ou URL do primeiro site estão inválidas, tente novamente!" -ForegroundColor Red
        do {
            $retry = Read-Host 'Qual a url do site de onde a lista será copiada?'; 
            ##NAO UTILIZADO AINDA
            $retry2 = Read-Host 'Qual a url do site para qual a lista será criada?';

            $res = retryConnection -siteurl $retry -isExternal $true;
            if ($res -eq "UriFormatException") {
                Write-Host "Url não válida!" -ForegroundColor Red      
            }
            if ($res -eq "WebException") {
                Write-Host "As credenciais ou URL do primeiro site estão inválidas, tente novamente!" -ForegroundColor Red      
            }
            if ($res -eq "IdcrlException") {
                Write-Host "As credenciais estão inválidas, tente novamente!" -ForegroundColor Red 
            }
            if ($res -eq "MultipleConnection") {
                #Executa funcao de pegar lista no mesmo tenanat
                $ListDe = Read-Host 'Qual lista deseja copiar?';
                if ($ListDe -ne $null) {
                    $ListDe = "Lists/" + $ListDe;        
                }
                $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
                if ($null -ne $ListPara -and $null -ne $ListDe) {
                    #Executa a copyListandCreate
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $Site2;
                    if ($res -eq "Valor inválido") {
            
                        do {      
                            Write-Host "Listas não encontradas, insira novamente!" -ForegroundColor Red
                            $ListDe = Read-Host 'Qual lista deseja copiar?';
                            if ($ListDe -ne $null) {
                                $ListDe = "Lists/" + $ListDe;        
                            }
                            $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
                            if ($ListPara -ne $null) {
                                $ListPara = "Lists/" + $ListPara;
                            }
                            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $retry2;
                        }
                        while ($res -eq "Valor inválido") 
                    };

                }
            }
            if ($res -eq "SingleConnection") {
                #Executa funcao de pegar lista em outro tenanat
                $ListDe = Read-Host 'Qual lista deseja copiar?';
                if ($ListDe -ne $null) {
                    $ListDe = "Lists/" + $ListDe;        
                }
                $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
                if ($null -ne $ListPara -and $null -ne $ListDe) {
                    #Executa a copyListandCreate
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara;
                    if ($res -eq "Valor inválido") {
            
                        do {      
                            Write-Host "Listas não encontradas, insira novamente!" -ForegroundColor Red
                            $ListDe = Read-Host 'Qual lista deseja copiar?';
                            if ($ListDe -ne $null) {
                                $ListDe = "Lists/" + $ListDe;        
                            }
                            $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
                            if ($ListPara -ne $null) {
                                $ListPara = "Lists/" + $ListPara;
                            }
                            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara;
                        }
                        while ($res -eq "Valor inválido") 
                    };

                }
            }
        }
        while ($res -eq "UriFormatException" -or $res -eq "WebException" -or $res -eq "IdcrlException") 
    };
    #Se for em mais um tenanat
    elseif ($result -eq "MultipleConnection") {
        #Executa funcao de pegar lista no mesmo tenanat
        $ListDe = Read-Host 'Qual lista deseja copiar?';
        if ($ListDe -ne $null) {
            $ListDe = "Lists/" + $ListDe;        
        }
        $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
        if ($null -ne $ListPara -and $null -ne $ListDe) {
            #Executa a copyListandCreate
            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $Site2;
            if ($res -eq "Valor inválido") {
    
                do {      
                    Write-Host "Listas não encontradas, insira novamente!" -ForegroundColor Red
                    $ListDe = Read-Host 'Qual lista deseja copiar?';
                    if ($ListDe -ne $null) {
                        $ListDe = "Lists/" + $ListDe;        
                    }
                    $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
                    if ($ListPara -ne $null) {
                        $ListPara = "Lists/" + $ListPara;
                    }
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $Site2;
                }
                while ($res -eq "Valor inválido") 
            };

        }
    }
    #Se for apenas no mesmo tenanat
    elseif ($result -eq "SingleConnection") {
        #Executa funcao de pegar lista em outro tenanat
        $ListDe = Read-Host 'Qual lista deseja copiar?';
        if ($ListDe -ne $null) {
            $ListDe = "Lists/" + $ListDe;        
        }
        $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
        if ($null -ne $ListPara -and $null -ne $ListDe) {
            #Executa a copyListandCreate
            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara;
            if ($res -eq "Valor inválido") {
    
                do {      
                    Write-Host "Listas não encontradas, insira novamente!" -ForegroundColor Red
                    $ListDe = Read-Host 'Qual lista deseja copiar?';
                    if ($ListDe -ne $null) {
                        $ListDe = "Lists/" + $ListDe;        
                    }
                    $ListPara = Read-Host 'Qual o nome da lista que deseja criar??';
                    if ($ListPara -ne $null) {
                        $ListPara = "Lists/" + $ListPara;
                    }
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara;
                }
                while ($res -eq "Valor inválido") 
            };

        }
    }
}