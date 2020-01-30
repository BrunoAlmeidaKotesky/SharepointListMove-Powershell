#Global
$currentTime = $(get-date).ToString("yyyyMMddHHmmss")  
$logFilePath = ".\log-" + $currentTime + ".docx"  

function loadLists {
    param([string]$ListDe, [string]$ListPara)
    #Pegando as listas
   
    if ($ListDe -eq $null -or $ListDe -eq "") {
        return "Valor inválido";
    }

    if ($ListPara -eq $null -or $ListPara -eq "") {
        return "Valor inválido";
    }

    $sourceList = Get-PnPList -Identity $ListDe;
    $targetList = Get-PnPList -Identity $ListPara;
    if ($sourceList -eq $null -or $targetList -eq $null -or $ListPara -eq $null -or $ListDe -eq $null) {
        return "Valor inválido";
    } 
    else {

        [array]$sourceItems = Get-PnPListItem -List $ListDe;
        #Array de colunas source e target
        $sourceFields = $sourceList.Fields
        $targetFields = $targetList.Fields
        #Carregando o contexto da lista
        $sourceList.Context.Load($sourceFields);
        $sourceList.Context.ExecuteQuery();
        $targetList.Context.Load($targetFields);
        $targetList.Context.ExecuteQuery();
        #Listas de campos
        $listaEncontrados = @()
        $listaNaoEncontrados = @()
        #Para cada coluna nas colunas da source
        foreach ($Field in $sourceFields | Where-Object { $_.FromBaseType -eq $false }) {
            #No Array de colunas do target, procure onde o nome interno for igual ao nome interno da source e tenha sido criado por usuário
            $itemEncontrado = $targetFields | Where-Object { $Field.FromBaseType -eq $false -and $_.FromBaseType -eq $false };
            $itemEncontrado = $targetFields | Where-Object { $_.InternalName -in $Field.InternalName }
            if ($itemEncontrado -ne $Null) {
                $listaEncontrados += $itemEncontrado[0].InternalName
            } 
            else {
                $listaNaoEncontrados += $Field.InternalName
            }
        }
        #No array de items da source, para cada item, criar um json vazio, e adicionando os campos
        foreach ($item in $sourceItems) {
            $jsonBase = @{"Title" = $item["Title"]; "Modified" = $item["Modified"]; "Created" = $item["Created"]; }
            #Para cada campo na lista de campos encontrados, adicione em um json
            $identifyTitle = Get-PnPListItem -List $ListPara -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$($item["Title"])</Value></Eq></Where></Query></View>";
            foreach ($campo in $listaEncontrados) {
                $jsonBase.Add($campo, $item[$campo]);
            }
            if ($identifyTitle.Length -gt 0) {
                #Adicione cada item com os valores do json montado
                Set-PnPListItem -List $ListPara -Values $jsonBase -Identity $identifyTitle.Id;
            }
            else {
                Add-PnPListItem -List $ListPara -Values $jsonBase
            }
        }
        try {
            $outputFilePath = ".\results-" + $currentTime + ".csv";
            $hashTable = @();
            foreach ($campo in $listaNaoEncontrados) {  
                $obj = New-Object PSObject              
                $obj | Add-Member -MemberType NoteProperty -name "CamposNaoEncontrados" -value $campo;
                $obj | Add-Member -MemberType NoteProperty -name "CamposAlvo" -value " ";
                $hashTable += $obj;  
                $obj = $null;  
                Write-Host "Colunas não preenchidas de $($ListDe) para a $($ListPara):", $campo -ForegroundColor Blue
            }
            $hashtable | Export-Csv $outputFilePath -NoTypeInformation;
            $targetFieldsEncontrados = $targetFields | Where-Object { $_.FromBaseType -eq $false };
            foreach ($coluna in $targetFieldsEncontrados) {
                Write-Host "Coluna presente na lista $($ListPara):", $coluna.InternalName -ForegroundColor Yellow
            }
            Write-Host "Um arquivo csv contendo as colunas da $($ListDe) não presentes na $($ListPara) foi criado com sucesso!" -ForegroundColor Green
            Start-Sleep -s 2
            Write-Host "Items enviados para a lista com sucesso!" -ForegroundColor Green
            Start-Sleep -s 6

        }
        catch [Exception] {  
            $ErrorMessage = $_.Exception.Message         
            Write-Host "Error: $ErrorMessage" -ForegroundColor Red          
        } 
        Stop-Transcript
    }
}
function tryToConnect {
    Param ([string]$siteurl, [bool]$isExternal = $false, [string]$siteurl2 = $null)
    if($isExternal -eq $false -or !$isExternal){
    try {
        Connect-PnPOnline -Url $siteurl -Credentials (Get-Credential);
        
        Start-Transcript -Path $logFilePath   
        #Definindo inputs
        $ListDe = Read-Host 'Qual lista deseja copiar?';
        if ($ListDe -ne $null) {
            $ListDe = "Lists/" + $ListDe;        
        }
        $ListPara = Read-Host 'Para qual lista deseja enviar?';
        if ($ListPara -ne $null) {
            $ListPara = "Lists/" + $ListPara;
        }

        #Pegando as listas
        $res = loadLists -ListDe $ListDe -ListPara $ListPara;
        if ($res -eq "Valor inválido") {
            
            do {      
                Write-Host "Listas não encontradas, insira novamente!" -ForegroundColor Red
                $ListDe = Read-Host 'Qual lista deseja copiar?';
                if ($ListDe -ne $null) {
                    $ListDe = "Lists/" + $ListDe;        
                }
                $ListPara = Read-Host 'Para qual lista deseja enviar?';
                if ($ListPara -ne $null) {
                    $ListPara = "Lists/" + $ListPara;
                }
                $res = loadLists -ListDe $ListDe -ListPara $ListPara;
            }
            while ($res -eq "Valor inválido") 
        };
    }
    catch [Exception] {  
        $ErrorMessage = $_.CategoryInfo.Reason; 
        return $ErrorMessage;
    }
 }
}
  

Write-host "A segunda lista está em outro ambiente do sharepoint? (Padrão: Não)" -ForegroundColor Yellow 
$userChoice = Read-Host " ( S / N ) "
Switch ($userChoice) { 
    S {
        $Site = Read-Host 'Qual a url do site de onde a lista é enviada?';
        #While pra validar se url é vazia
        while (!$Site) {
            $Site = Read-Host 'Qual a url do site de onde a lista é enviada?';
        }
        $Site2 = Read-Host 'Qual a url do site para qual a lista será enviada?';
        #While pra validar se url é vazia
        while (!$Site2) {
            $Site2 = Read-Host 'Qual a url do site para qual a lista será enviada?';
        }

    } 
    N {
        $Site = Read-Host 'Qual a url quer navegar?';
        #While pra validar se url é vazia
        while (!$Site) {
            $Site = Read-Host 'Qual a url quer navegar?';  
        }

        $result = tryToConnect -siteurl $Site;
        if ($result -eq "UriFormatException" -or $result -eq "WebException" -or $result -eq "IdcrlException") {
    
            do {
                $retry = Read-Host 'Qual a url quer navegar?'; 
                $res = tryToConnect -siteurl $retry
                if ($res -eq "UriFormatException") {
                    Write-Host "Url não válida!" -ForegroundColor Red      
                }
                if ($res -eq "WebException") {
                    Write-Host "As credenciais ou URL estão inválidas" -ForegroundColor Red      
                }
                if ($res -eq "IdcrlException") {
                    Write-Host "As credenciais estão inválidas" -ForegroundColor Red 
                }
            }
            while ($res -eq "UriFormatException" -or $res -eq "WebException" -or $res -eq "IdcrlException") 
        };
    } 
    Default {
        $Site = Read-Host 'Qual a url quer navegar?';
        #While pra validar se url é vazia
        while (!$Site) {
            $Site = Read-Host 'Qual a url quer navegar?';  
        }

        $result = tryToConnect -siteurl $Site;
        if ($result -eq "UriFormatException" -or $result -eq "WebException" -or $result -eq "IdcrlException") {
    
            do {
                $retry = Read-Host 'Qual a url quer navegar?'; 
                $res = tryToConnect -siteurl $retry
                if ($res -eq "UriFormatException") {
                    Write-Host "Url não válida!" -ForegroundColor Red      
                }
                if ($res -eq "WebException") {
                    Write-Host "As credenciais ou URL estão inválidas" -ForegroundColor Red      
                }
                if ($res -eq "IdcrlException") {
                    Write-Host "As credenciais estão inválidas" -ForegroundColor Red 
                }
            }
            while ($res -eq "UriFormatException" -or $res -eq "WebException" -or $res -eq "IdcrlException") 
        };
    } 
}

