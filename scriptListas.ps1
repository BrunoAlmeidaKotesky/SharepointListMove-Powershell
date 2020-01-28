$Site = Read-Host 'Qual a url quer navegar?';

while (!$Site) {
   $Site = Read-Host 'Adicione uma URL!'   
}

Connect-PnPOnline -Url $Site;

$ListDe = Read-Host 'Qual lista deseja copiar?';
$ListDe = "Lists/" + $ListDe;
$ListPara = Read-Host 'Para qual lista deseja enviar?';
$ListPara = "Lists/" + $ListPara;

$sourceList = Get-PnPList -Identity $ListDe;
$targetList = Get-PnPList -Identity $ListPara;    
$sourceItems = Get-PnPListItem -List $ListDe;

$sourceFields = $sourceList.Fields
$targetFields = $targetList.Fields

$sourceList.Context.Load($sourceFields);
$sourceList.Context.ExecuteQuery();
$targetList.Context.Load($targetFields);
$targetList.Context.ExecuteQuery();



   $listaEncontrados = @()
   $listaNaoEncontrados = @()

   foreach ($Field in $sourceFields) {

    $itemEncontrado = $targetFields | Where-Object { $_.InternalName -in $Field.InternalName -and $Field.FromBaseType -eq $false -and $_.FromBaseType -eq $false }
    if($itemEncontrado -ne $null)
    {
        $listaEncontrados += $itemEncontrado[0].InternalName
    } else
    {
        $listaNaoEncontrados += $Field.InternalName
    }
   }
   

   foreach ($item in $sourceItems) 
    {
        $jsonBase = @{}

        foreach($campo in $listaEncontrados)
        {
           $jsonBase.Add($campo, $item[$campo]); 
        }
        
       Add-PnPListItem -List $ListPara -Values $jsonBase

        #$jsonBase
    }
 
 # criar lista de campos
 # criar lista de campos não encontrados
 # iteração para cada field do Source
    # Valida se o campo existe no target, usando um filtro
    # SE o campo existe
        # adiciona o nome interno do campo na lista de campos
    # SENAO
        # adiciona o nome interno do campo na lista de campos nao encontrados
# iteração para cada item da lista de source
# criar objeto json
# iteração para cada campo da lista de campos encontrados
    # crair uma proprierdade no objeto json
# rodar add-pnplistitem com o valor do json gerado
