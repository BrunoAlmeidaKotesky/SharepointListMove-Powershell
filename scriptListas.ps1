Connect-PnPOnline -Url https://trentim.sharepoint.com -UseWebLogin

$ListDe = Read-Host 'Qual lista deseja copiar?'
$ListDe = "Lists/" + $ListDe;
$ListPara = Read-Host 'Para qual lista deseja enviar?'
$ListPara = "Lists/" + $ListPara

try {
    $targetItems = Get-PnPListItem -List $ListPara
    $sourceItems = Get-PnPListItem -List $ListDe
    $sourceList = Get-PnPList -Identity $ListDe
    $targetList = Get-PnPList -Identity $ListPara

    $sourceList.Context.Load($sourceList.Fields);
    $sourceList.Context.ExecuteQuery();
    $targetList.Context.Load($targetList.Fields);
    $targetList.Context.ExecuteQuery();

    $itemsIguaisSource = (Get-PnPListItem -List $ListDe -Fields "Title","GUID")
    $itemsIguaisTarget = (Get-PnPListItem -List $ListPara -Fields "Title","GUID")

    foreach ($Field in $sourceList.Fields) {
        foreach ($Field2 in $targetList.Fields) {
            if ($Field.FromBaseType -eq $false -and $Field2.FromBaseType -eq $false ) {
                Write-Host "O campo '" $Field.Title "' foi criado pelo usuário(Nome interno é '"$Field.InternalName"')";
                Write-Host "O campo '" $Field2.Title "' foi criado pelo usuário(Nome interno é '"$Field2.InternalName"')";
                if ($Field.InternalName -eq $Field2.InternalName) {
                    Write-Host "São iguais"
                    Add-PnPListItem -List $ListPara -Identity $Field2['ID'] -Values @{"MultiUserField"="user1@domain.com","user2@domain.com"}

                }
                else {
                    Write-Host $Field.InternalName
                }
            }
        }
    }
}
catch {
    Write-Host "Ocorreu um erro"
}