function Stop-ServiceWithTimeout ([string] $name, [int] $timeoutSeconds) #������� ��������� ������
{
		$timespan = New-Object -TypeName System.Timespan -ArgumentList 0,0,$timeoutSeconds
		$svc = Get-Service -Name $name
		if ($svc -eq $null) 
		{ 
			return $false 
		}
		if ($svc.Status -eq [ServiceProcess.ServiceControllerStatus]::Stopped) 
		{ 
			return $true 
		}
		$svc.Stop()
		try {
			$svc.WaitForStatus([ServiceProcess.ServiceControllerStatus]::Stopped, $timespan)
		}
		catch [ServiceProcess.TimeoutException] {
			Write-Verbose "Timeout stopping service $($svc.Name)"
			return $false
		}
	
    return $true
}

function roboError
{
	Switch ($LASTEXITCODE)
	{
		16
		{
			$exit_code = "16"
			$exit_reason = "��������� ������! ������ �� �����������, ��������� ������, ������ � ������ ��� ����� �������"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		15
		{
			$exit_code = "15"
			$exit_reason = "[FAILED] OKCOPY + FAIL MISMATCH EXTRA COPY"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		14
		{
			$exit_code = "14"
			$exit_reason = "[FAILED] FAIL MISMATCH EXTRA"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		13
		{
			$exit_code = "13"
			$exit_reason = "[FAILED] OKCOPY + FAIL MISMATCH COPY"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		12
		{
			$exit_code = "12"
			$exit_reason = "������ ��������������"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		11
		{
			$exit_code = "11"
			$exit_reason = "[FAILED] OKCOPY + FAIL EXTRA COPY"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		10
		{
			$exit_code = "10"
			$exit_reason = "������"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		9
		{
			$exit_code = "9"
			$exit_reason = "������ �����������"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		8
		{
			$exit_code = "8"
			$exit_reason = "[����������� �����] �� ������� ����������� ��������� ����� ��� ��������, � ��� �������� ������ ��������� �������"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		7
		{
			$exit_code = "7"
			$exit_reason = "����� ���� �����������, �������������� ����� ��������������, � �������������� ����� ��������������"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		6
		{
			$exit_code = "6"
			$exit_reason = "���������� �������������� ����� � ����������������� �����. ����� �� ���� ����������� � ����� �� ����. ��� ��������, ��� ����� ��� ���������� � �������� ����������"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		5
		{
			$exit_code = "5"
			$exit_reason = "��������� ����� ���� �����������. ��������� ����� �� �������. ����� �� ����������"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		4
		{
			$exit_code = "4"
			$exit_reason = "���������� ������������� ����� ��� ��������. �������� ������"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		3
		{
			$exit_code = "3"
			$exit_reason = "��������� ����� ���� �����������. �������������� ����� ��������������. ����� �� ����������"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='green'>OK</font>"
		}
		2
		{
			$exit_code = "2"
			$exit_reason = "�������������� ����� ��� �������� ���� ����������. �������� ������"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='green'>OK</font>"
		}
		1
		{
			$exit_code = "1"
			$exit_reason = "���� ��� ��������� ������ ���� �����������"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='green'>OK</font>"
		}
		0
		{
			$exit_code = "0"
			$exit_reason = "����� �� �����������, ����� �� ����������"
			$backupState = "<font size='3' color='green'>OK</font>"
			$SendEmail = $False
			$IncludeAdmin = $False
		}
		default
		{
			$exit_code = "Unknown ($LASTEXITCODE)"
			$exit_reason = "Unknown Reason"
			$IncludeAdmin = $False
		}
	}
		[string]$date = get-date -Format "dd.MM.yyyy HH:mm:ss";
	$ret = "<td>"+$exit_reason+"</td><td>"+$backupState+"</td><td>"+$date+"</td></tr>"
	Return $ret
	

}

function checkDir([string]$path)
{
	if (!(test-path -path $path)) 
	{
		new-item -path $path -itemtype directory
	}
}

# ������� ���� �� �����, �� exe ����� �� �������� $PSScriptRoot
$path_to_script = 
	if (-not $PSScriptRoot) 
	{  
		Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0])) 
	} 
	else 
	{
		$PSScriptRoot 
	}

# �������� ������ �� JSON
$json = Get-Content ( $path_to_script + '\ps_backup.json')  
$json = $json -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*' -replace '(?ms)/\*.*?\*/' 
$json = $json | Out-String | ConvertFrom-Json



$json.name_dir   

$splitter = "`n_____________________________________________________________________________________________`n`n"




$path_to_backup_dir = $json.path_to_logs

# ������������� � ��������� ����� � ������
$path_to_logs = $path_to_backup_dir + 'logs/'
checkDir($path_to_logs)

# ������������� � ��������� ����� � ��������
$path_to_archives = $path_to_backup_dir  + 'archives/'
checkDir($path_to_archives)

[string]$fileDayLog = ($path_to_logs+"DAY.log")
[string]$fileAllLog = ($path_to_logs+"ALL.log")
[string]$fileHtml = ($path_to_logs+"day.html")
[string]$pathArch = ($path_to_logs+"logs-all.zip")

$fileDayLog
$fileAllLog
$fileHtml 
$pathArch 
#Get-Process powershell | Out-File $fileDayLog -append

#$date | Out-File $fileDayLog -append

#������ ��������� ��� robocopy

#��� ����������� �����	
$name_dir = $json.name_dir
$source = $json.source
$destination = $json.destination

#��� ������������� (��������� ���������� �� ������� �����)
# ������� ���� ������� ������
$day_save_arc = $json.day_save_arc

#������������ ������
$name_arc_rus = $json.name_arc_rus
	
$list_arc =  @(	#������ ������, ������� �� ������� ��������
	($destination[0])
	#,($destination[1])
)

$destination_arc = $json.destination_arc 

#������� � ����� ������
$name_arc = $json.name_arc  

#��������� ��������� ��� ��������	
$robocopyOptions = @(
'/Z' #����������� ������ � ������ �����������
,'/MIR' #��������� �������� ������ �����
,'/R:5' #��������� ����� ��������� ������� ��� ����������� �����
,'/W:15' #��������� ����� �������� ����� ���������� ��������� � ��������
,'/V' #������� ��������� �������� ������ � ���������� ��� ����������� �����
,'/TS' #�������� � ���� ������� ������� ��������� ����� � �������� ������ 
#,'/FP' #�������� ������ ���� ����� ������ � �������� ������
#,'/NP' #���������, ��� ��� ���������� �������� (����� ������ ��� ���������, ���������� � ������ ������) �� ����� ������������
,'/TEE' #���������� �������� ������ � ���������, � ���� �������, � ����� � ���� �������
,'/NDL' #��� ������ �����
,'/NFL' #���������, ��� ����� ������ �� ��� ��������� � ������
#,'/NJH' #���������, ��� ����������� ��������� �������.
);
[array]$strLog = $strBegin;	
[string]$date = get-date -Format "dd.MM.yyyy HH:mm:ss";
$strLog = $splitter + "**********START "+$date+"**********"+$splitter #| Out-File $fileDayLog;
[string]$sendMailText = "<center><h2> ����� �� ������ �� "+$date + "</h1></center>"
$sendMailText += "<center><table border='1'   cellspacing='0' cellpadding = '5' >
  <tr >
	<td>������������</td>
    <td>���� ��</td>
    <td>���� �</td>
    <td>���� ������</td>
    <td>���������</td>
    <td>������</td>
    <td>���� ���������</td>
   </tr>"
for ($i = 0; $i -lt $destination.length; $i++)
{ 
	[string]$date = get-date -Format "dd.MM.yyyy HH:mm:ss";
	$sendMailText += "<tr><td>"+$name_dir[$i]+"</td><td>"+$source[$i] + "</td><td>"+$destination[$i]+"</td><td>"+$date+"</td>"
	$CmdLine = @($source[$i], $destination[$i], $fileList) + $robocopyOptions;

	#�������� ����� �� �������
	checkDir($destination[$i])
	
	#��������� RoboCopy � ��������� �����������
	Try	{
		$strLog +=&'robocopy.exe' $CmdLine
	}
	Catch	{
		$strLog +="������ ��� ������� RoboCopy: "+$error[0].Exception
	}

	$tempVal = roboError
	$sendMailText += $tempVal;
	"����������� ��������� "+$name_dir[$i]+" - "+$tempVal
}
$sendMailText += "</table></center>"

#��������� �������
#$sendMailText +="<br><br><br><center><h2> ����� �� ������� </h2></center>"
$sendMailText += "<center><table border='1'   cellspacing='0' cellpadding = '5' >
  <tr >
	<td>������������</td>
    <td>������������� ��</td>
    <td>������������� �</td>
    <td>���������</td>
    <td>������</td>
    <td>������������</td>
	<td>���������</td>
	<td>������</td>
   </tr>"
for ($i = 0; $i -lt $list_arc.length; $i++)
{
	checkDir($destination_arc[$i]);
	[string]$date = get-date -Format "yyyy-MM-dd";
	$fullDirArc = $destination_arc[$i]+$name_arc[$i]+$date+".zip"

	#��������� �������: ������������, ������������� ��, ������������� �
	$sendMailText += "<tr><td>"+$name_arc_rus[$i]+"</td><td>"+$list_arc[$i]+"</td><td>"+$fullDirArc+"</td>"
	Try{
		Compress-Archive -Path $list_arc[$i] -DestinationPath $fullDirArc -CompressionLevel Optimal -Force
		$tempVal ="����� ������� ������ "+$fullDirArc;
		$strLog +=$tempVal
		$arcStatus = "<td><font size='3' color='green'>OK</font></td>"
		}
	Catch{
		$tempVal ="������ ��� �������������: "+$fullDirArc+ $error[0].Exception
		$strLog += $tempVal
		$arcStatus = "<td><font size='3' color='red'>ERROR</font></td>"
		}
	$sendMailText += "<td>"+$tempVal+"</td>"+$arcStatus
	"������������� ��������� "+$name_arc_rus[$i]+" - "+$tempVal
	
	#������� ������ ������ �������� ����
	$sendMailText += "<td>�������� �������</td>"
	#"�������� ������� " + $dateDelArc 
	
	#�������� ������� ���� - ���������� ���� ��������� ��� �������� �������
	$dateDelArc = (get-date).addDays($day_save_arc[$i])
	#�������� ������ �������, ��� ���� ������ ����� ������ ���� ��������� ��� �������� �������
	$list_old_arc = Get-Item ($destination_arc[$i]+"*.zip") | where lastWriteTime -lt $dateDelArc
	
	#��������� ����� �� �������������
	if ($list_old_arc -ne $NULL)
	{
		Try{
			#$list_old_arc | Remove-Item
			$tempVal = "������� ������: "+$list_old_arc
			$strLog += $tempVal;
			$arcStatus = "<td><font size='3' color='green'>OK</font></td>"
			}
		Catch{
			$tempVal ="������ ��� �������� �������: "+$error[0].Exception
			$strLog +=$tempVal
			$arcStatus = "<td><font size='3' color='red'>ERROR</font></td>"
			}
	}
	else{
		$tempVal ="��� ������� ��� ��������: "+$error[0].Exception
		$strLog += $tempVal;
		$arcStatus = "<td><font size='3' color='green'>OK</font></td>"
		}
		
		$sendMailText += "<td>$tempVal</td>"+$arcStatus+"</tr>"
}

$sendMailText += "</table></center>"

$sendMailText| Out-File $fileHtml
[string]$date = get-date -Format "yyyy.MM.dd HH:mm:ss";
$strLog += $splitter + "**********END "+$date+"**********"+$splitter;

$strLog| Out-File $fileDayLog 
$strLog| Out-File $fileAllLog -append
#out-string -inputobject $strLog
#$strLog = get-content -Path $fileDayLog
#$strLog
#$strLog |  Out-File  $fileDayLog #| out-string
#$fileList = $fileList + "12323"
#$fileList
#pause 2000

#$PSScriptRoot #���� �� �������


#�������� ����� 
#�������� ������ ���������:
# $From = "fgdg@fgfg.com"
# $To = "fgdg@fgfg.com"
# $SMTPServer = "smtp.gmail.com"
# $SMTPPort = "587"
# $Username = "fgdg@fgfg.com"
# $Password = "pass"
# $subject = "����� �� ���������� �����������"
# $body = $sendMailText;

# #��������� ��������� � ������� html:
# $message = New-Object System.Net.Mail.MailMessage $From, $To
# $message.Subject = $subject
# $message.IsBodyHTML = $true
# $message.Body = $body

# #����������:

# $smtp = New-Object System.Net.Mail.SmtpClient($SMTPServer, $SMTPPort)
# $smtp.EnableSSL = $true
# $smtp.Credentials = New-Object System.Net.NetworkCredential($Username, $Password)
#$smtp.Send($message)
#$message


<#$service = "spooler"; 
if ((Get-Service $service).Status -eq 'Stopped') 
{
	Start-Service $service
}#>




