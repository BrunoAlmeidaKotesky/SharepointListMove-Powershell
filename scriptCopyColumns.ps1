
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


function addFields() {
    param($sourceFields, [string]$ListPara, $ctx, [bool]$isExternal);
    foreach ($field in $sourceFields | Where-Object { $_.FromBaseType -eq $false }) {
        if($field.InternalName -ne "Title" -or $field.InternalName -ne "Modified" -or $field.InternalName -ne "Created"){
            if($field.Required -eq $true){
                if($field.FieldTypeKind -eq "Lookup"){
                    $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Required -Type Lookup  -InternalName $field.InternalName;
                    $lkField = $novaColuna.TypedObject;
                    $lookId1 = $field.LookupList.Replace("{", "");
                    $lookId2 = $lookId1.Replace("}", "");
                    if($true -eq $isExternal){
                        
                    }
                    else {
                        $lkField.LookupList = $lookId2;  #use the actual ID of the list, not the name
                        $lkField.LookupField = $field.LookupField;
                    }
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
                    if($true -eq $isExternal){
                        $ls = Get-PnPList -Identity $ListPara;
                        $lkField.LookupList = $ls.Id  #use the actual ID of the list, not the name
                        $lkField.LookupField = $field.LookupField;
                    }
                    else {
                        $lkField.LookupList = $lookId2;  #use the actual ID of the list, not the name
                        $lkField.LookupField = $field.LookupField;
                    }
                    $lkField.update();
                    $ctx.ExecuteQuery();
                }
                else{
                    $novaColuna = Add-PnPField -List $ListPara -AddToDefaultView -DisplayName $field.Title -Type $field.TypeAsString  -InternalName $field.InternalName;
                }
                
            }
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
                        catch [Excption]{
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
                catch [Excption]{
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
