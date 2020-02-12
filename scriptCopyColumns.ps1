
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

function userOptions {
    param([bool]$option = $true)
    $Site = Read-Host 'Qual a url do site de onde a lista sera copiada?';
    #While pra validar se url é vazia
    while (!$Site) {
        $Site = Read-Host 'Qual a url do site de onde a lista sera copiada?';
    }
    $connectionType;

    if($false -eq $option){
        $Site2 = $null;
        $connectionType = $false;
    }
    else{
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
                    elseif($res -eq "TenantDisconnected"){

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
            elseif ($res -eq "TenantDisconnected"){

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

function insertAllColumns (){
    param($allCols, $ListPara);

    foreach ($field in $allCols) {
        if($field.InternalName -ne "Title" -or $field.InternalName -ne "Modified" -or $field.InternalName -ne "Created"){
            if($field.Required -eq $true){
                if($field.FieldTypeKind -eq "Lookup"){
                    $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Required -Type Lookup  -InternalName $field.InternalName;
                    $lkField = $novaColuna.TypedObject;
                    $lookId1 = $field.LookupList.Replace("{", "");
                    $lookId2 = $lookId1.Replace("}", "");
                    $lkField.LookupList = $lookId2;  #use the actual ID of the list, not the name
                    $lkField.LookupField = $field.LookupField;
                    $lkField.update();
                    $ctx.ExecuteQuery();
                }
                else {
                    $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Required -Type $field.TypeAsString  -InternalName $field.InternalName;
                }
            }
            else{
                if($field.FieldTypeKind -eq "Lookup"){
                    $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Type Lookup -InternalName $field.InternalName;
                    $lkField = $novaColuna.TypedObject;
                    $lookId1 = $field.LookupList.Replace("{", "");
                    $lookId2 = $lookId1.Replace("}", "");
                    $lkField.LookupList = $lookId2;  #use the actual ID of the list, not the name
                    $lkField.LookupField = $field.LookupField;
                    $lkField.update();
                    $ctx.ExecuteQuery();
                }
                else{
                    $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Type $field.TypeAsString  -InternalName $field.InternalName;
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

    if($true -eq $isExternal){
      $colunasComLookup += $sourceFields | ? {$_.FieldTypeKind -eq "Lookup"};
      if($colunasComLookup.Count -gt 0) {
         Write-Host "Na lista origem ha colunas do tipo lookup, deseja ignorar esses campos ou especificar qual lista e campo sera enviada para o site alvo?" -ForegroundColor Yellow;
         
           $option = Read-Host "[I](Ignorar)/[E](Especificar)"

          Switch($option)
            {
             E {
                 $newUserFields += $sourceFields | Where-Object { $_.FieldTypeKind -ne "Lookup" };
                 foreach($newField in $colunasComLookup){
                     if($newField.FieldTypeKind -eq "Lookup"){
                         $listName = Read-Host "Qual e lista para o $($newField.InternalName)";
                         $colName = Read-Host "Qual e coluna para o $($newField.InternalName)";
                         $ls = Get-PnPList -Identity $listName;
                         $lookObject | Add-Member -type NoteProperty -name listId -Value $ls.Id;
                         $lookObject | Add-Member -type NoteProperty -name colName -Value $colName;
                         $lookObject | Add-Member -type NoteProperty -name title -Value $newField.Title;
                         $lookObject | Add-Member -type NoteProperty -name internalName -Value $newField.InternalName;
                     }
                 }
                 $lookObject | ForEach-Object {
                     $newCol = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $_.title-Type Lookup -InternalName $_.internalName; 
                     $lkField = $newCol.TypedObject;
                     $lkField.LookupList = $_.listId;  #use the actual ID of the list, not the name
                     $lkField.LookupField = $_.colName;
                     $lkField.update();
                     $ctx.ExecuteQuery();
                 } 
                 insertAllColumns -allCols $newUserFields -ListPara $ListPara;
             }
             I { 
                  $newUserFields += $sourceFields | Where-Object { $_.FieldTypeKind -ne "Lookup" };
                  insertAllColumns -allCols $newUserFields -ListPara $ListPara;
             }
        }
      }
      else { insertAllColumns -allCols $sourceFields -ListPara $ListPara;}
    }
    else{
         $newUserFields = $sourceFields;
         insertAllColumns -allCols $newUserFields -ListPara $ListPara;
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
        if($true -eq $lostContext){
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
                    if($res -eq "MultipleConnection"){
                        $sourceList = Get-PnPList -Identity $ListDe;
                        $allSourceFields =Get-PnPField -List $ListDe
                        $ctx = Get-PnPContext;
                        $sourceFields = $allSourceFields | Where-Object { $_.FromBaseType -eq $false };
                    }
                } while($res -eq "UriFormatException" -or $res -eq "WebException" -or $res -eq "IdcrlException")
            }
            elseif ($tenant1 -eq "MultipleConnection"){
                $sourceList = Get-PnPList -Identity $ListDe;
                $allSourceFields =Get-PnPField -List $ListDe
                $ctx = Get-PnPContext;
                $sourceFields = $allSourceFields | Where-Object { $_.FromBaseType -eq $false };
            }
        }
        #Carregando o contexto da lista
        $sourceList = Get-PnPList -Identity $ListDe;
        $allSourceFields =Get-PnPField -List $ListDe
        $ctx = Get-PnPContext;
        $sourceFields = $allSourceFields | Where-Object { $_.FromBaseType -eq $false };
        #Se for no mesmo tenant
        if ($null -eq $segundoSite -or $segundoSite -eq "") {
            if ($null -eq $sourceList) { return "Valor inválido"; } 
            $listExists = Get-PnPList -Identity $ListPara.Replace("Lists/", "");
            if($null -eq $listExists){
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
            else{ Write-Host "Ja existe uma lista com esse nome!" -ForegroundColor Red
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
                    if($res -eq "MultipleConnection"){
                        $listExists = Get-PnPList -Identity $ListPara.Replace("Lists/", "");
                        if($null -eq $listExists){
                            $novaLista = New-PnPList -Title $ListPara.Replace("Lists/", "") -Template GenericList;
                        try {
                            addFields -sourceFields $sourceFields -ListPara $ListPara -ctx $ctx -isExternal $true;
                            Write-Host "Lista criada com sucesso!" -ForegroundColor Green;  
                        }
                        catch [Exception]{
                            Remove-PnPList -Identity $ListPara -Force;
                            return $_;
                        }
                        }
                        else{ Write-Host "Ja existe uma lista com esse nome!" -ForegroundColor Red
                              return "ListAlreadyExists";
                        }
                    }
                }
                while ($res -eq "UriFormatException" -or $res -eq "WebException" -or $res -eq "IdcrlException") 
            }
            elseif ($connectionRes -eq "MultipleConnection") {
                $listExists = Get-PnPList -Identity $ListPara.Replace("Lists/", "");
                if($null -eq $listExists){
                    $novaLista = New-PnPList -Title $ListPara.Replace("Lists/", "") -Template GenericList;
                try {
                    addFields -sourceFields $sourceFields -ListPara $ListPara -ctx $ctx -isExternal $true;
                    Write-Host "Lista criada com sucesso!" -ForegroundColor Green;  
                }
                catch [Exception]{
                    Remove-PnPList -Identity $ListPara -Force;
                    return $_;
                }
                }
                else{ Write-Host "Ja existe uma lista com esse nome!" -ForegroundColor Red
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
