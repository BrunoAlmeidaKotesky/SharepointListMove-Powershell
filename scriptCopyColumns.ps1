
function retryConnection {
    Param ([string]$siteurl , [boolean]$isExternal = $false)

    try {
        if ($true -eq $isExternal -and $null -ne $siteurl) {
            #Add-PnPStoredCredential -Name $siteurl;             
            Write-Host "Insira as credenciais da URL!" -ForegroundColor Yellow
            Connect-PnPOnline -Url $siteurl;
            return "MultipleConnection";
        }
        elseif ($false -eq $isExternal -or $null -eq $isExternal) {
            Write-Host "Insira as credenciais da URL!" -ForegroundColor Yellow
            Connect-PnPOnline -Url $siteurl;
            return "SingleConnection";
        }
    }
    catch [Exception] {
        $ErrorMessage = $_.CategoryInfo.Reason; 
        return $ErrorMessage;
    }
}

class LookUpCol {
    [string]$listId
    [string]$colName
    [string]$title
    [string]$internalName
    [bool]$isMulti
}

function userOptions {
    param([bool]$option = $true)
    $Site = Read-Host 'Qual a url do site de onde a lista sera copiada?';
    #While pra validar se url é vazia
    while (!$Site) {
        $Site = Read-Host 'Qual a url do site de onde a lista sera copiada?';
    }
    $connectionType;

    if ($false -eq $option) {
        $Site2 = $null;
        $connectionType = $false;
    }
    else {
        $Site2 = Read-Host 'Qual a url do site para qual a lista sera criada?';
        #While pra validar se url é vazia
        while (!$Site2) {
            $Site2 = Read-Host 'Qual a url do site para qual a lista sera criada?';
        }
        $connectionType = $true;
    }

    $result = retryConnection -siteurl $Site -isExternal $connectionType;
    if ($result -eq "UriFormatException" -or $result -eq "WebException" -or $result -eq "IdcrlException") {
        Write-Host "As credenciais ou URL do primeiro site estao invalidas, tente novamente!" -ForegroundColor Red
        do {
            $retry = Read-Host 'Qual a url do site de onde a lista sera copiada?'; 
            ##NAO UTILIZADO AINDA
            $retry2 = Read-Host 'Qual a url do site para qual a lista sera criada?';

            $res = retryConnection -siteurl $retry -isExternal $connectionType;
            if ($res -eq "UriFormatException") {
                Write-Host "Url nao valida!" -ForegroundColor Red      
            }
            if ($res -eq "WebException") {
                Write-Host "As credenciais ou URL do primeiro site estao invalidas, tente novamente!" -ForegroundColor Red      
            }
            if ($res -eq "IdcrlException") {
                Write-Host "As credenciais estao invalidas, tente novamente!" -ForegroundColor Red 
            }
            if ($res -eq "MultipleConnection") {
                #Executa funcao de pegar lista no mesmo tenanat
                $ListDe = Read-Host 'Qual lista deseja copiar?';
                if ($ListDe -ne $null) {
                    $ListDe = "Lists/" + $ListDe;        
                }
                $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
                if ($null -ne $ListPara -and $null -ne $ListDe) {
                    #Executa a copyListandCreate
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $Site2;
                    if ($res -eq "ListAlreadyExists") {
            
                        do {      
                            Write-Host "Listas nao encontradas, insira novamente!" -ForegroundColor Red
                            $ListDe = Read-Host 'Qual lista deseja copiar?';
                            if ($ListDe -ne $null) {
                                $ListDe = "Lists/" + $ListDe;        
                            }
                            $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
                            if ($ListPara -ne $null) {
                                $ListPara = "Lists/" + $ListPara;
                            }
                            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $retry2;
                        }
                        while ($res -eq "ListAlreadyExists" -or $res -eq "Valor Inválido") 
                    }
                    elseif ($res -eq "TenantDisconnected") {

                        do {      
                            Write-Host "Listas nao encontradas, insira novamente!" -ForegroundColor Red
                            $ListDe = Read-Host 'Qual lista deseja copiar?';
                            if ($ListDe -ne $null) {
                                $ListDe = "Lists/" + $ListDe;        
                            }
                            $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
                            if ($ListPara -ne $null) {
                                $ListPara = "Lists/" + $ListPara;
                            }
                            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -firstUrl $retry -segundoSite $retry2 -lostContext $true;
                        }
                        while ($res -eq "TenantDisconnected")
                    }

                }
            }
            if ($res -eq "SingleConnection") {
                #Executa funcao de pegar lista em outro tenanat
                $ListDe = Read-Host 'Qual lista deseja copiar?';
                if ($ListDe -ne $null) {
                    $ListDe = "Lists/" + $ListDe;        
                }
                $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
                if ($null -ne $ListPara -and $null -ne $ListDe) {
                    #Executa a copyListandCreate
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara;
                    if ($res -eq "Valor inválido") {
            
                        do {      
                            Write-Host "Listas nao encontradas, insira novamente!" -ForegroundColor Red
                            $ListDe = Read-Host 'Qual lista deseja copiar?';
                            if ($ListDe -ne $null) {
                                $ListDe = "Lists/" + $ListDe;        
                            }
                            $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
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
    }
    elseif ($result -eq "MultipleConnection") {
        #Executa funcao de pegar lista no mesmo tenanat
        $ListDe = Read-Host 'Qual lista deseja copiar?';
        if ($ListDe -ne $null) {
            $ListDe = "Lists/" + $ListDe;        
        }
        $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
        if ($null -ne $ListPara -and $null -ne $ListDe) {
            #Executa a copyListandCreate
            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $Site2;
            if ($res -eq "ListAlreadyExists") {
                
                do {      
                    Write-Host "Listas nao encontradas, insira novamente!" -ForegroundColor Red
                    $ListDe = Read-Host 'Qual lista deseja copiar?';
                    if ($ListDe -ne $null) {
                        $ListDe = "Lists/" + $ListDe;        
                    }
                    $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
                    if ($ListPara -ne $null) {
                        $ListPara = "Lists/" + $ListPara;
                    }
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -segundoSite $Site2;
                }
                while ($res -eq "ListAlreadyExists" -or $res -eq "Valor Inválido") 
            }
            elseif ($res -eq "TenantDisconnected") {

                do {      
                    Write-Host "Listas nao encontradas, insira novamente!" -ForegroundColor Red
                    $ListDe = Read-Host 'Qual lista deseja copiar?';
                    if ($ListDe -ne $null) {
                        $ListDe = "Lists/" + $ListDe;        
                    }
                    $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
                    if ($ListPara -ne $null) {
                        $ListPara = "Lists/" + $ListPara;
                    }
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara -firstUrl $Site -segundoSite $Site2 -lostContext $true;
                }
                while ($res -eq "TenantDisconnected")
            }

        }
    }
    #Se for apenas no mesmo tenanat
    elseif ($result -eq "SingleConnection") {
        #Executa funcao de pegar lista em outro tenanat
        $ListDe = Read-Host 'Qual lista deseja copiar?';
        if ($ListDe -ne $null) {
            $ListDe = "Lists/" + $ListDe;        
        }
        $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
        if ($null -ne $ListPara -and $null -ne $ListDe) {
            #Executa a copyListandCreate
            $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara;
            if ($res -eq "Valor inválido") {
    
                do {      
                    Write-Host "Listas nao encontradas, insira novamente!" -ForegroundColor Red
                    $ListDe = Read-Host 'Qual lista deseja copiar?';
                    if ($ListDe -ne $null) {
                        $ListDe = "Lists/" + $ListDe;        
                    }
                    $ListPara = Read-Host 'Qual o nome da lista que deseja criar?';
                    if ($ListPara -ne $null) {
                        $ListPara = "Lists/" + $ListPara;
                    }
                    $res = copyAndCreateList -ListDe $ListDe -ListPara $ListPara;
                }
                while ($res -eq "Valor inválido") 
            }
        }
    }
}

function insertAllColumns () {
    param($allCols, $ListPara);
    foreach ($field in $allCols) {
        Write-Progress -Id 1  -Activity Updating -Status "Criando colunas: Isso pode demorar alguns minutos!" -PercentComplete (($count / $allCols.Count) * 100);
        if ($field.InternalName -ne "Title" -or $field.InternalName -ne "Modified" -or $field.InternalName -ne "Created") {
            Switch($field.Required){
                $true{
                    Switch($field.FieldTypeKind){
                        Choice{
                            if ($field.TypedObject.EditFormat -eq "RadioButtons") {
                                $choices = @();
                                foreach ($choice in $field.Choices) {
                                    $choice = $choice.Replace("/", "");
                                    $choice = $choice.Replace("&", "");
                                    $choice = $choice.Replace("(", "");
                                    $choice = $choice.Replace(")", "");
                                    $choices += "<CHOICE>$($choice)</CHOICE>"
                                };
                            
                                $radioBtnField = "<Field Type='Choice' ShowInViewForms='TRUE' DisplayName='$($field.Title)' Format='RadioButtons' Required='TRUE' StaticName='$($field.InternalName)' Name='$($field.InternalName)'>
                                              <CHOICES>$($choices)</CHOICES></Field>";
                                $novaColunaXMLChoices = Add-PnPFieldFromXml -List $ListPara -FieldXml $radioBtnField;
                            }
                            else { $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Required -Type Choice -Choices $field.Choices -InternalName $field.InternalName; }
                        }
                        MultiChoice{
                            $choices = @();
                            foreach ($choice in $field.Choices) {
                                $choice = $choice.Replace("/", "");
                                $choice = $choice.Replace("&", "");
                                $choice = $choice.Replace("(", "");
                                $choice = $choice.Replace(")", "");
                                $choices += "<CHOICE>$($choice)</CHOICE>"
                            };
                            
                            $chkBoxField = "<Field Type='MultiChoice' ShowInViewForms='TRUE' DisplayName='$($field.Title)' Required='TRUE' StaticName='$($field.InternalName)' Name='$($field.InternalName)'>
                                              <CHOICES>$($choices)</CHOICES></Field>";
                            $novaColunaXMLChoices = Add-PnPFieldFromXml -List $ListPara -FieldXml $chkBoxField;
                        }
                        DateTime{

                        }
                        Default{ $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Required -Type $field.TypeAsString  -InternalName $field.InternalName;}
                    }
                }
                $false{
                    Switch ($field.FieldTypeKind) {
                        Choice {
                            if ($field.TypedObject.EditFormat -eq "RadioButtons") {
                                $choices = @();
                                foreach ($choice in $field.Choices) {
                                    $choice = $choice.Replace("/", "");
                                    $choice = $choice.Replace("&", "");
                                    $choice = $choice.Replace("(", "");
                                    $choice = $choice.Replace(")", "");
                                    $choices += "<CHOICE>$($choice)</CHOICE>"
                                };
                            
                                $radioBtnField = "<Field Type='Choice' ShowInViewForms='TRUE' DisplayName='$($field.Title)' Format='RadioButtons' Required='FALSE' StaticName='$($field.InternalName)' Name='$($field.InternalName)'>
                                              <CHOICES>$($choices)</CHOICES></Field>";
                                $novaColunaXMLChoices = Add-PnPFieldFromXml -List $ListPara -FieldXml $radioBtnField;
                            }
                            else { $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Type Choice -Choices $field.Choices -InternalName $field.InternalName; }
                        }
    
                        MultiChoice {
                            $choices = @();
                            foreach ($choice in $field.Choices) {
                                $choice = $choice.Replace("/", "");
                                $choice = $choice.Replace("&", "");
                                $choice = $choice.Replace("(", "");
                                $choice = $choice.Replace(")", "");
                                $choices += "<CHOICE>$($choice)</CHOICE>"
                            };
                            
                            $chkBoxField = "<Field Type='MultiChoice' ShowInViewForms='TRUE' DisplayName='$($field.Title)' Required='FALSE' StaticName='$($field.InternalName)' Name='$($field.InternalName)'><CHOICES>$($choices)</CHOICES></Field>";
                            $novaColunaXMLChoices = Add-PnPFieldFromXml -List $ListPara -FieldXml $chkBoxField;
                        }
                        DateTime {
                            
                        }
                        Default { $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Type $field.TypeAsString  -InternalName $field.InternalName; }
                    }
                }
            }
        }
    }
}

function addFields() {
    param($sourceFields, [string]$ListPara, $ctx, [bool]$isExternal);
    $newUserFields = @();
    $colunasComLookup = @();
    $lookObject = New-Object System.Object;
    #Colunas que possuem lookup
    $colunasComLookup += $sourceFields | ? { $_.FieldTypeKind -eq "Lookup" -or $_.FieldTypeKind -eq "MultiLookup" };

    if ($true -eq $isExternal) {
        if ($colunasComLookup.Count -gt 0) {
            Write-Host "Na lista origem ha colunas do tipo lookup, deseja ignorar esses campos ou especificar qual lista e campo sera enviada para o site alvo?" -ForegroundColor Yellow;
         
            $option = Read-Host "[I](Ignorar)/[E](Especificar)"

            Switch ($option) {
                E {
                    $newUserFields += $sourceFields | Where-Object { $_.FieldTypeKind -ne "Lookup" };
                    $lookObject = @();
                    foreach ($newField in $colunasComLookup) {
                        if ($newField.TypeAsString -eq "Lookup") {
                            $listName = Read-Host "Qual e lista para o $($newField.InternalName)";
                            $colName = Read-Host "Qual e coluna para o $($newField.InternalName)";
                            $ls = Get-PnPList -Identity $listName;
                            $lookObject += @([LookUpCol]@{
                                    listId       = $ls.Id;
                                    colName      = $colName;
                                    title        = $newField.Title;
                                    internalName = $newField.InternalName;
                                    isMulti      = $false;
                                })
                        }
                        elseif ($newField.TypeAsString -eq "LookupMulti") {
                            $listName = Read-Host "Qual e lista para o $($newField.InternalName)";
                            $colName = Read-Host "Qual e coluna para o $($newField.InternalName)";
                            $ls = Get-PnPList -Identity $listName;
                            $lookObject += @([LookUpCol]@{
                                    listId       = $ls.Id;
                                    colName      = $colName;
                                    title        = $newField.Title;
                                    internalName = $newField.InternalName;
                                    isMulti      = $true;
                                })
                        }
                    }
                    foreach ($item in $lookObject) {
                        if ($item.isMulti -eq $false) {
                            $newCol = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $item.title-Type Lookup -InternalName $item.internalName; 
                            $lkField = $newCol.TypedObject;
                            $lkField.LookupList = $item.listId; #use the actual ID of the list, not the name
                            $lkField.LookupField = $item.colName;
                            $lkField.update();
                            $ctx.ExecuteQuery();
                        }
                        else {
                            $multiLookUp = "<Field Type='LookupMulti' List='$($item.listId)' DisplayName='$($item.title)' Required='FALSE' Mult='TRUE' EnforceUniqueValues='FALSE' ShowField='$($item.colName)' UnlimitedLengthInDocumentLibrary='FALSE' RelationshipDeleteBehavior='None' ShowInViewForms='TRUE' StaticName='$($item.internalName)' Name='$($item.internalName)'/>";
                            $novaColunaXMLChoices = Add-PnPFieldFromXml -List $ListPara -FieldXml $multiLookUp;
                        }
                    } 
                    insertAllColumns -allCols $newUserFields -ListPara $ListPara;
                }
                I { 
                    $newUserFields += $sourceFields | Where-Object { $_.FieldTypeKind -ne "Lookup" };
                    insertAllColumns -allCols $newUserFields -ListPara $ListPara;
                }
            }
        }
        #Se mais de um tenant E nao existir lookup insere normalmente
        else { insertAllColumns -allCols $sourceFields -ListPara $ListPara; }
    }
    #Se for somente um tenanat
    else {
        #se em um tenanat tiver lookup
        if ($colunasComLookup.Count -gt 0) {
            $newUserFields += $sourceFields | Where-Object { $_.FieldTypeKind -ne "Lookup" }; #Items que não sejam lookup nem multi lookup
            $lookObject = @();
            foreach ($item in $colunasComLookup) {
                if ($item.TypeAsString -eq "Lookup") {
                    $newCol = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $item.title-Type Lookup -InternalName $item.internalName; 
                    $lkField = $newCol.TypedObject;
                    $lookId1 = $item.LookupList.Replace("{", "");
                    $lookId2 = $lookId1.Replace("}", "");
                    $lkField.LookupList = $lookId2; #use the actual ID of the list, not the name
                    $lkField.LookupField = $item.LookupField;
                    $lkField.update();
                    $ctx.ExecuteQuery();
                }
                elseif ($item.TypeAsString -eq "LookupMulti") {
                    $lookId1 = $item.LookupList.Replace("{", "");
                    $lookId2 = $lookId1.Replace("}", "");
                    $multiLookUp = "<Field Type='LookupMulti' List='$($lookId2)' DisplayName='$($item.Title)' Required='FALSE' Mult='TRUE' EnforceUniqueValues='FALSE' ShowField='$($item.LookupField)' UnlimitedLengthInDocumentLibrary='FALSE' RelationshipDeleteBehavior='None' ShowInViewForms='TRUE' StaticName='$($item.InternalName)' Name='$($item.InternalName)'/>";
                    $novaColunaXMLChoices = Add-PnPFieldFromXml -List $ListPara -FieldXml $multiLookUp;
                }
            } 
            insertAllColumns -allCols $newUserFields -ListPara $ListPara;
        }
        #Se em um tenant nao tiver lookup
        else {
            insertAllColumns -allCols $sourceFields -ListPara $ListPara;
        }
    }
};

function copyAndCreateList {
    param([string]$ListDe, [string]$ListPara, [string]$segundoSite, [bool]$lostContext, [string]$firstUrl)
    #Verificando se o valor da lista é nulo
    if ($ListDe -eq $null -or $ListDe -eq "") {
        return "Valor inválido";
    }

    if ($ListPara -eq $null -or $ListPara -eq "") {
        return "Valor inválido";
    }
    else {
        if ($true -eq $lostContext) {
            Write-Host "Ocorreu um erro ao obter a primeira lista, por favor insira a primeira URL novamente para re-obter a lista" -ForegroundColor Yellow;
            $tenant1 = retryConnection -siteurl $firstUrl -isExternal $true;
            if ($tenant1 -eq "UriFormatException" -or $connectionRes -eq "WebException" -or $connectionRes -eq "IdcrlException") {
                Write-Host "A url ou credenciais informadas para o site estao invalidas, tente novamente!" -ForegroundColor Red;
                do {
                    $retry = Read-Host 'Insira novamente a primeira url relacionada a lista origem.';
    
                    $res = retryConnection -siteurl $retry -isExternal $true;
                    if ($res -eq "UriFormatException") {
                        Write-Host "Url nao valida!" -ForegroundColor Red      
                    }
                    if ($res -eq "WebException") {
                        Write-Host "As credenciais ou URL estão invalidas" -ForegroundColor Red      
                    }
                    if ($res -eq "IdcrlException") {
                        Write-Host "As credenciais estao invalidas" -ForegroundColor Red 
                    }
                    if ($res -eq "MultipleConnection") {
                        $sourceList = Get-PnPList -Identity $ListDe;
                        $allSourceFields = Get-PnPField -List $ListDe
                        $ctx = Get-PnPContext;
                        $sourceFields = $allSourceFields | Where-Object { $_.FromBaseType -eq $false };
                    }
                } while ($res -eq "UriFormatException" -or $res -eq "WebException" -or $res -eq "IdcrlException")
            }
            elseif ($tenant1 -eq "MultipleConnection") {
                $sourceList = Get-PnPList -Identity $ListDe;
                $allSourceFields = Get-PnPField -List $ListDe
                $ctx = Get-PnPContext;
                $sourceFields = $allSourceFields | Where-Object { $_.FromBaseType -eq $false };
            }
        }
        #Carregando o contexto da lista
        $sourceList = Get-PnPList -Identity $ListDe;
        $allSourceFields = Get-PnPField -List $ListDe
        $ctx = Get-PnPContext;
        $sourceFields = $allSourceFields | Where-Object { $_.FromBaseType -eq $false };
        #Se for no mesmo tenant
        if ($null -eq $segundoSite -or $segundoSite -eq "") {
            if ($null -eq $sourceList) { return "Valor inválido"; } 
            $listExists = Get-PnPList -Identity $ListPara.Replace("Lists/", "");
            if ($null -eq $listExists) {
                $novaLista = New-PnPList -Title $ListPara.Replace("Lists/", "") -Template GenericList;
                try {
                    addFields -sourceFields $sourceFields -ListPara $ListPara -ctx $ctx -isExternal $false;
                    Write-Host "Lista criada com sucesso!" -ForegroundColor Green;  
                }
                catch [Exception] {
                    Remove-PnPList -Identity $ListPara -Force;
                    return $_;
                }
            }
            else {
                Write-Host "Ja existe uma lista com esse nome!" -ForegroundColor Red
                return "Valor inválido";
            }
        }
        #Se for em mais de um tenanat
        else {
            if ($null -eq $sourceList) { return "TenantDisconnected"; } 
            Disconnect-PnPOnline;
            Write-Host "Insira as credenciais da segunda URL para onde a lista sera copiada" -ForegroundColor Yellow;
            $connectionRes = retryConnection -siteurl $segundoSite -isExternal $true;
            if ($connectionRes -eq "UriFormatException" -or $connectionRes -eq "WebException" -or $connectionRes -eq "IdcrlException") {
                Write-Host "A url ou credenciais informadas para o site alvo estão invalidas, tente novamente!" -ForegroundColor Red;
                do {
                    $retry = Read-Host 'Qual a url do site para qual a lista sera criada?';
    
                    $res = retryConnection -siteurl $retry -isExternal $true;
                    if ($res -eq "UriFormatException") {
                        Write-Host "Url nao valida!" -ForegroundColor Red      
                    }
                    if ($res -eq "WebException") {
                        Write-Host "As credenciais ou URL estao invalidas" -ForegroundColor Red      
                    }
                    if ($res -eq "IdcrlException") {
                        Write-Host "As credenciais estao invalidas" -ForegroundColor Red 
                    }
                    if ($res -eq "MultipleConnection") {
                        $listExists = Get-PnPList -Identity $ListPara.Replace("Lists/", "");
                        if ($null -eq $listExists) {
                            $novaLista = New-PnPList -Title $ListPara.Replace("Lists/", "") -Template GenericList;
                            try {
                                addFields -sourceFields $sourceFields -ListPara $ListPara -ctx $ctx -isExternal $true;
                                Write-Host "Lista criada com sucesso!" -ForegroundColor Green;  
                            }
                            catch [Exception] {
                                Remove-PnPList -Identity $ListPara -Force;
                                return $_;
                            }
                        }
                        else {
                            Write-Host "Ja existe uma lista com esse nome!" -ForegroundColor Red
                            return "ListAlreadyExists";
                        }
                    }
                }
                while ($res -eq "UriFormatException" -or $res -eq "WebException" -or $res -eq "IdcrlException") 
            }
            elseif ($connectionRes -eq "MultipleConnection") {
                $listExists = Get-PnPList -Identity $ListPara.Replace("Lists/", "");
                if ($null -eq $listExists) {
                    $novaLista = New-PnPList -Title $ListPara.Replace("Lists/", "") -Template GenericList;
                    try {
                        addFields -sourceFields $sourceFields -ListPara $ListPara -ctx $ctx -isExternal $true;
                        Write-Host "Lista criada com sucesso!" -ForegroundColor Green;  
                    }
                    catch [Exception] {
                        Remove-PnPList -Identity $ListPara -Force;
                        return $_;
                    }
                }
                else {
                    Write-Host "Ja existe uma lista com esse nome!" -ForegroundColor Red
                    return "ListAlreadyExists";
                }
            }
        }
    }
}

Write-host "A transferencia da lista, ira ocorrer no em outro ambiente do Sharepoint?" -ForegroundColor Yellow 
$userChoice = Read-Host " ( S / N ) "
Switch ($userChoice) { 
    S { userOptions -option $true; }
    N { userOptions -option $false; } 
    Default { userOptions -option $true; }
}
