#Defines

<#
Function get-admins
Devuelve los usuarios de dominio de una máquina pasada como parámetro
#>
function get-admins{
param ($strcomputer)
    $admins = Gwmi win32_groupuser -computer $strcomputer
    $admins = $admins |? {$_.partcomponent -like '*UserAccount.Domain="COMPANY"*'} #La palabra COMPANY se debe sustituir por el nombre de dominio de las cuentas de usuario
    $ret= @($admins |% {  
        $_.partcomponent –match “.+Domain\=(.+)\,Name\=(.+)$” > $nul  
        $matches[2].trim('"')
    })
    $ret
}

<#
Function check_format
Comprueba que el formato de máquina pasado como parámetro se corresponda con un nombre o una IP
#>
function check_format{
param($strcomputer)
    $ip_format = "^[0-9]{1,3}.[0-9]{1,3}.[0-9]{1,3}.[0-9]{1,3}$"
    $machine_format = "^[a-zA-Z0-9]{1,20}$"
    if(($computer -match $ip_format) -bor ($computer -match $machine_format)){
        $true
    }else{
        $false
    }
}

#Execution
<#
-Se comprueba para cada uno de los parámetros de entrada que el formato sea correcto.
-Si es correcto se prueba a obtener los usuarios de dominio de la misma.
-Se ordenan los resultados en un diccionario en orden de último uso.
-Se presentan los resultados.
#>
$file_path = $args[0]

for($i=0; $i -lt $args.Length; $i++){
    $computer = $args[$i]

    #Comprobación de los parámetros de entrada.
    if(check_format($computer)){
        $maquina_str = $computer
        $maquina_str_out = "MAQUINA: " + $computer
        Write-Output $maquina_str_out
        Write-Output ""
        #Obtención de los adminsitradores de la máquina
        $admins = @(get-admins($computer))
        $hashtable = @{}
        foreach($item in $admins){
            $hashtable[$item] = $(Get-ChildItem "\\$computer\c$\Users" | Where-Object Name -Like $item | Select-Object LastWriteTime).LastWriteTime
        }
        $hashtable_o = [ordered]@{}

        $maquina_str | Out-File -FilePath .\arch.txt -Append -NoNewline
        #Ordenación de los resultados
        foreach($item in $hashtable.GetEnumerator() | Sort Value -Descending){
            $hashtable_o[$item.Name] = $item.Value
        }
        
        #Presentación de los resultados en formato tabla con fecha de último acceso
        <#Write-Output ""
        $hashtable_o
        Write-Output ""#>

        #Presentación de los resultados en formato String ordenada de usuarios con acceso másreciente
        $str_result = "Los últimos usuarios que han accedido son: "
        $keys = @($hashtable_o.keys)
        for($k = 0; $k -lt $hashtable_o.count; $k++){
            if($k -eq $hashtable_o.count - 1) {
                $str_result= $str_result + $keys[$k]
            }else{
                $str_result= $str_result + $keys[$k] + ", "
            }
        }
        [char]9 + $str_result | Out-File -FilePath .\arch.txt -Append
        Write-Output $str_result
        Write-Output ""
        Write-Output "---------------"
    #Parámetros de entrada incorrectos
    }else{
        $maquina_str = $computer
        $maquina_str_out = "MAQUINA: " + $computer
        Write-Output $maquina_str_out
        Write-Output ""
        $maquina_str | Out-File -FilePath .\arch.txt -Append -NoNewline
        Write-Output "El formato introducido no es correcto"
        Write-Output ""
        [char]9 + "El formato introducido no es correcto" | Out-File -FilePath .\arch.txt -Append
        Write-Output "---------------"
    }
}

#Formato de archivos de salida
$date = Get-Date -Format "dd_MM_yyyy-HHmm"
Copy-Item .\arch.txt -Destination .\Equipos-$date.txt
Remove-Item .\arch.txt
$exitstring = "Archivo de salida: " + ".\Equipos-" + $date + ".txt"
Write-Output ""
Write-Output $exitstring
Write-Output "" 