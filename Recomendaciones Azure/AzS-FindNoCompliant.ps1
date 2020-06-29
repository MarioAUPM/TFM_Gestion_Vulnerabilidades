$ErrorActionPreference = 'Stop'
$RecommendationTable = @()
$RecommendationTableNOTRES = 
$MissingSubscriptions = @()
$Subscriptions = Get-AzSubscription
$inputFolder = "" + $(Get-Location)

<#
Para cada suscripción, extrae las recomendaciones y las guarda en formato tabla con los campos:
    -Recommendation
    -Resource
    -Subscription Name
    -Subscription Id
    -Ressource Group
    -Tennant (Aun vacío)
Y almacena esta información en la tabla RecommendationsTable
En caso de no poder extraer recomendaciones para la suscripción, la saltará y continuará.
#>
$counter_g = 0
$correctos_g = 0
$fallos_g = 0
$Sin_resource_g = 0
foreach($Subscription in $Subscriptions){
    Select-AzSubscription $Subscription
    $SecurityTasks = Get-AzSecurityTask

    try{
    $counter = 0
    $correctos = 0
    $fallos = 0
    $Sin_resource = 0
    foreach($SecurityTask in $SecurityTasks){
            $counter++
            $counter_g++
            If([string]::IsNullOrEmpty($SecurityTask.ResourceId.Split("/")[8])) {  
                $RecommendationsNOTRES = New-Object psobject -Property @{
                    Recommendation = $SecurityTask.RecommendationType
                    Resource = ($SecurityTask.ResourceId.Split("/")[8])
                    SubscriptionName = $Subscription.Name
                    SubscriptionId = ($SecurityTask.ResourceId.Split("/")[2])
                    ResourceGroup = ($SecurityTask.ResourceId.Split("/")[4])
                    Tennant = ""
                }
                $RecommendationTableNOTRES += $RecommendationsNOTRES
                $Sin_resource++
                $Sin_resource_g++
            }
            else {
                $Recommendations = New-Object psobject -Property @{
                    Recommendation = $SecurityTask.RecommendationType
                    Resource = ($SecurityTask.ResourceId.Split("/")[8])
                    SubscriptionName = $Subscription.Name
                    SubscriptionId = ($SecurityTask.ResourceId.Split("/")[2])
                    ResourceGroup = ($SecurityTask.ResourceId.Split("/")[4])
                    Tennant = ""
                }
                $correctos++
                $correctos_g++
                $RecommendationTable += $Recommendations
            }
        }
    }
    catch
    {
        Write-Host "Could not get recommendations for subscription: " $Subscription.Name -ForegroundColor Red
        Write-Host "Error Message: " $_.Exception.Message -ForeGroundColor Red
        Write-Host "Skipping subscription `r`n" -ForegroundColor Red
        $MissingSubscriptionsDetails = New-Object psobject -Property @{
            SubscriptionName = $Subscription.Name
            SubscriptionId = ($SecurityTask.ResourceId.Split("/")[2])
            ErrorMessage = $_.Exception.Message
        }
        $MissingSubscriptions += $MissingSubscriptionsDetails
        $fallos++
    }
    $str = "recomendaciones: " + $counter + "; correctos: " + $correctos + "; sin recurso asociado: " + $Sin_resource + "; fallos: " + $fallos
    Write-Host $str -ForegroundColor Green

}

$str_g = "TOTAL => recomendaciones: " + $counter_g + "; correctos: " + $correctos_g + "; sin recurso asociado: " + $Sin_resource_g + "; fallos: " + $fallos_g
Write-Host $str_g -ForegroundColor Green

<#
Almacena esta información SIN RECURSO en un fichero CSV
#>
try
{
    Write-Host "Almacenando recomendaciones sin recursos asociados en CSV" -ForegroundColor Yellow    
    $name = ("tmp_sin_recurso.csv")
    $RecommendationTableNOTRES | Select-Object "SubscriptionName", "SubscriptionId", "Resource", "Recommendation", "ResourceGroup", "Tennant" | Export-Csv -Path ($name) -Force -NoTypeInformation
    Write-Host "Done! `r`n" -ForegroundColor Yellow
}
catch {Write-Host "Could not create output file.... Please check your path, filename and write permissions." -ForeGroundColor Red}

<#
Para cada recomendación, busca el contacto técnico de la máquina
#>
$Counter2 = 0
$Correctos2 = 0
$Fallos2 = 0
$Sin_responsable = 0
foreach($Recommendation in $RecommendationTable){
        $Counter2++
        try{
            $Resource = $Recommendation.ResourceGroup
            Write-Host $Resource -ForegroundColor Yellow
            $Subscription = Get-AzSubscription -SubscriptionId $Recommendation.SubscriptionId
            Select-AzSubscription $Subscription
            $Responsable = Get-AzureRmResourceGroup -Name ($Resource) |findstr "Technical"
            if([string]::IsNullorEmpty($Responsable)){
                $Sin_responsable++
                Write-Host "Sin responsable" -ForegroundColor Red
                $Recommendation.Tennant = "SIN RESPONSABLE"
            }else{
                $Responsable = $Responsable.REPLACE(" ","")
                $Responsable = $Responsable.REPLACE("TechnicalContact","")
                $Recommendation.Tennant = $Responsable
                $Correctos2++
                }
        }
        catch{
            $Fallos2++
            Write-Host "RESOURCE GROUP NO EXISTE" -ForegroundColor Red
            $Recommendation.Tennant = "NO EXISTE RECURSO"
        }
}
$str2 = "responsables: " + $Counter2 + "; correctos: " + $Correctos2 + "; sin responsable: " + $Sin_responsable + "; fallos: " + $Fallos2
Write-Host $str2 -ForegroundColor Green

<#
Almacena esta información en un fichero CSV
#>
try
{
    Write-Host "Almacenando recomendaciones de recursos en CSV" -ForegroundColor Yellow   

    #$date = Get-Date -Format "dd_MM_yyyy-HHmm"
    $name = "tmp_con_recurso.csv"
    $RecommendationTable | Select-Object "SubscriptionName", "SubscriptionId", "Resource", "Recommendation", "ResourceGroup", "Tennant" | Export-Csv -Path ($name) -Force -NoTypeInformation
    Write-Host "Done! `r`n" -ForegroundColor Yellow
}
catch {Write-Host "Could not create output file.... Please check your path, filename and write permissions." -ForeGroundColor Red}



<#
---CONVERSION DE LOS FICHEROS CSV TEMPORALES A EXCEL---
#>



### Set input and output path
$inputCSV = $inputFolder + "\tmp_con_recurso.csv"
$inputCSV2 = $inputFolder + "\tmp_sin_recurso.csv"
$date = Get-Date -Format "dd_MM_yyyy-HHmm"
$outputXLSX = $inputFolder + "\AzureSecurityCenter_Recommendations-" + $date + ".xlsx"

### Create a new Excel Workbook with one empty sheet
$excel = New-Object -ComObject excel.application
$excel.sheetsInNewWorkbook = 5 
$workbooks = $excel.Workbooks.Add()
$worksheets = $workbooks.worksheets
$worksheet1 = $worksheets.Item(1)
$worksheet2 = $worksheets.Item(2)
$worksheet3 = $worksheets.Item(3)
$worksheet4 = $worksheets.Item(4)
$worksheet5 = $worksheets.Item(5)
$arr = $worksheet1, $worksheet2, $worksheet3, $worksheet4, $worksheet5

for($i=0; $i -lt 4; $i++){
    $sheet = $arr[$i]
    Switch($i){
     0{$sheet.Name = "ASC-Recommendations"}
     1{$sheet.Name = "Por Responsable"}
     2{$sheet.Name = "Por Recomendacion"}
     3{$sheet.Name = "Por Suscripcion"}
    }
    ### Build the QueryTables.Add command
    ### QueryTables does the same as when clicking "Data » From Text" in Excel
    $TxtConnector = ("TEXT;" + $inputCSV)
    $Connector = $sheet.QueryTables.add($TxtConnector,$sheet.Range("A1"))
    $query = $sheet.QueryTables.item($Connector.name)

    ### Set the delimiter (, or ;) according to your regional settings
    $query.TextFileOtherDelimiter = ','

    ### Set the format to delimited and text for every column
    ### A trick to create an array of 2s is used with the preceding comma
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,2 * $sheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1

    ### Execute & delete the import query
    $query.Refresh()
    $query.Delete()
}

##Ordering Worksheets
$xlSortOnValues = 0
$xlTopToBottom  = 1
$xlAscending    = 1
$xlDescending   = 2
$xlNo           = 2
$xlYes          = 1

#POR RESPONSABLE

    $objRange = $worksheet2.UsedRange
    $objRange1 = $worksheet2.Range("F1")
    $worksheet2.Sort.SortFields.Clear()

    [void] $worksheet2.Sort.SortFields.Add($objRange1,$xlSortOnValues,$xlAscending,$xlSortNormal)

    $worksheet2.sort.setRange($objRange)
    $worksheet2.sort.header = $xlYes
    $worksheet2.sort.apply()

#POR RECOMENDACIÓN

    $objRange = $worksheet3.UsedRange
    $objRange1 = $worksheet3.Range("D1")
    $worksheet3.Sort.SortFields.Clear()

    [void] $worksheet3.Sort.SortFields.Add($objRange1,$xlSortOnValues,$xlAscending,$xlSortNormal)

    $worksheet3.sort.setRange($objRange)
    $worksheet3.sort.header = $xlYes
    $worksheet3.sort.apply()

#POR SUBSCRIPCIÓN

    $objRange = $worksheet4.UsedRange
    $objRange1 = $worksheet4.Range("A1")
    $worksheet4.Sort.SortFields.Clear()

    [void] $worksheet4.Sort.SortFields.Add($objRange1,$xlSortOnValues,$xlAscending,$xlSortNormal)

    $worksheet4.sort.setRange($objRange)
    $worksheet4.sort.header = $xlYes
    $worksheet4.sort.apply()

#RECOMENDACIONES DE SUSCRIPCION
    $sheet = $arr[4]
    $sheet.Name = "Recomendaciones de suscripcion"

    ### Build the QueryTables.Add command
    ### QueryTables does the same as when clicking "Data » From Text" in Excel
    $TxtConnector = ("TEXT;" + $inputCSV2)
    $Connector = $sheet.QueryTables.add($TxtConnector,$sheet.Range("A1"))
    $query = $sheet.QueryTables.item($Connector.name)

    ### Set the delimiter (, or ;) according to your regional settings
    $query.TextFileOtherDelimiter = ','

    ### Set the format to delimited and text for every column
    ### A trick to create an array of 2s is used with the preceding comma
    $query.TextFileParseType  = 1
    $query.TextFileColumnDataTypes = ,2 * $sheet.Cells.Columns.Count
    $query.AdjustColumnWidth = 1

    ### Execute & delete the import query
    $query.Refresh()
    $query.Delete()

### Save & close the Workbook as XLSX. Change the output extension for Excel 2003
$workbooks.SaveAs($outputXLSX,51)
$excel.Quit()


###Remove Temporal Items
Remove-Item $inputCSV
Remove-Item $inputCSV2
