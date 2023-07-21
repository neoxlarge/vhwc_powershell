$all_user = Get-WmiObject Win32_userAccount | Select-Object Name

$Ago = (Get-Date).Adddays(-3)

$log_event = Get-WinEvent -FilterHashtable @{LogName='Security'; ID='4624'; StartTime=$Ago} 


for ($i = 0; $i -lt $all_user.Count; $i++) {

    for ($j = 0; $j -lt $log_event.Count; $j++) {
        
        $result = $log_event[$j].properties[5].value -eq ($all_user[$i].name)


        if ($result) {
            #Write-Output $all_user[$i] "login"
            out-file -InputObject $all_user[$i] -FilePath remote_logined.log -Append
            break
        } 

    }

	if (!$result) {
	#write-output $all_user[$i] " not login"
    out-file -InputObject $all_user[$i] -FilePath remote_NOT_logined.log -Append
    }
}