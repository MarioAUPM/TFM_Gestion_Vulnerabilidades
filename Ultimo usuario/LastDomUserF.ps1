#Defines

<#
Function get-admins
Devuelve los usuarios de dominio de una máquina pasada como parámetro
Usa PSEXEC, por lo que este tiene que estar en la ruta C:\Program Files\PSTools\PsExec
#>
function get-admins{
param ($strcomputer)
    $users = &("C:\Program Files\PSTools\PsExec.exe") \\$strcomputer net localgroup administrators
    $ret = @(foreach($item in $users){
        if($item -like "COMPANY\*"){ #La palabra COMPANY se debe sustituir por el nombre de dominio de las cuentas de usuario
            $item_2 = $item.Replace("COMPANY\","") #La palabra COMPANY se debe sustituir por el nombre de dominio de las cuentas de usuario
            $item_2
        }
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
for($i=0; $i -lt $args.Length; $i++){
    $computer = $args[$i]
    

    #Comprobación de los parámetros de entrada.
    if(check_format($computer)){
        $maquina_str = $args[$i]
        $maquina_str_out = "MAQUINA: " + $args[$i]
        Write-Output $maquina_str_out
        Write-Output ""
        #Obtención de los adminsitradores de la máquina
        $admins = @(get-admins($computer))
        $hashtable = @{}
        foreach($item in $admins){
            $hashtable[$item] = $(Get-ChildItem "\\$computer\c$\Users" | Where-Object Name -Like $item | Select-Object LastWriteTime).LastWriteTime
        }
        $hashtable_o = [ordered]@{}

        #Ordenación de los resultados
        foreach($item in $hashtable.GetEnumerator() | Sort Value -Descending){
            $hashtable_o[$item.Name] = $item.Value
        }

        $maquina_str | Out-File -FilePath .\arch.txt -Append -NoNewline
        #Presentación de los resultados en formato String ordenada de usuarios con acceso másreciente a un archivo
        $str_result = "" + [char]9
        $keys = @($hashtable_o.keys)
        for($k = 0; $k -lt $hashtable_o.count; $k++){
            if($k -eq $hashtable_o.count - 1) {
                $str_result= $str_result + $keys[$k]
            }else{
                $str_result= $str_result + $keys[$k] + ", "
            }
        }
        $str_result | Out-File -FilePath .\arch.txt -Append
        Write-Output "---------------"

    #Parámetros de entrada incorrectos
    }else{
        $maquina_str = $args[$i]
        $maquina_str_out = "MAQUINA: " + $args[$i]
        Write-Output $maquina_str_out
        Write-Output ""
        Write-Output "El formato introducido no es correcto"
        Write-Output ""
        $maquina_str | Out-File -FilePath .\arch.txt -Append -NoNewline
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