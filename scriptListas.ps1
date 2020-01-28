$Site = Read-Host 'Qual a url quer navegar?';

while (!$Site) {
   $Site = Read-Host 'Adicione uma URL!'   
}

#Connect-PnPOnline -Url 'https://trentim.sharepoint.com';

$ListDe = Read-Host 'Qual lista deseja copiar?';
$ListDe = "Lists/" + 'ListaLivroGP';
$ListPara = Read-Host 'Para qual lista deseja enviar?';
$ListPara = "Lists/" +'testeDavi';

$sourceList = Get-PnPList -Identity $ListDe;
$targetList = Get-PnPList -Identity $ListPara;    
 $sourceItems = Get-PnPListItem -List $ListDe;

$sourceFields = $sourceList.Fields
$targetFields = $targetList.Fields

$sourceList.Context.Load($sourceFields);
$sourceList.Context.ExecuteQuery();
$targetList.Context.Load($targetFields);
$targetList.Context.ExecuteQuery();


   foreach ($Field in $sourceFields) {
      foreach ($Field2 in $targetFields) {
         if ($Field.FromBaseType -eq $false -and $Field2.FromBaseType -eq $false) {
            if ($Field.InternalName -eq $Field2.InternalName) {
               if ($Field.ReadOnlyField -eq $False -and $Field.InternalName -ne "Attachments" -and $Field2.ReadOnlyField -eq $False -and $Field2.InternalName -ne "Attachments") {
                 <#  Copy-PnPFile -SourceUrl $sourceList -TargetUrl $targetList #>
                   foreach ($item in $sourceItems) 
                   {
                     Add-PnPListItem -List $ListPara -Values @{"Title" = $item["Title"]; $Field.InternalName = $item[$Field.InternalName]}
                   }
               }
            }
         }
      }
   }

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
