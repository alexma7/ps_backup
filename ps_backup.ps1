function Stop-ServiceWithTimeout ([string] $name, [int] $timeoutSeconds) #Функция остановки службы
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
			$exit_reason = "ФАТАЛЬНАЯ ошибка! Ничего не скопировано, проверьте журнал, доступ к данным или права доступа"
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
			$exit_reason = "ОШИБКА несоответствия"
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
			$exit_reason = "ОШИБКА"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		9
		{
			$exit_code = "9"
			$exit_reason = "ОШИБКА копирования"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		8
		{
			$exit_code = "8"
			$exit_reason = "[НЕУДАЧЕННЫЕ КОПИИ] Не удалось скопировать некоторые файлы или каталоги, и был превышен предел повторных попыток"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		7
		{
			$exit_code = "7"
			$exit_reason = "Файлы были скопированы, несоответствие файла присутствовало, и дополнительные файлы присутствовали"
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		6
		{
			$exit_code = "6"
			$exit_reason = "Существуют дополнительные файлы и несоответствующие файлы. Файлы не были скопированы и сбоев не было. Это означает, что файлы уже существуют в каталоге назначения"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		5
		{
			$exit_code = "5"
			$exit_reason = "Некоторые файлы были скопированы. Некоторые файлы не совпали. Сбоев не обнаружено"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		4
		{
			$exit_code = "4"
			$exit_reason = "Обнаружены несовместимые файлы или каталоги. Смотрите журнал"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='red'>ERROR</font>"
		}
		3
		{
			$exit_code = "3"
			$exit_reason = "Некоторые файлы были скопированы. Дополнительные файлы присутствовали. Сбоев не обнаружено"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='green'>OK</font>"
		}
		2
		{
			$exit_code = "2"
			$exit_reason = "ДОПОЛНИТЕЛЬНЫЕ ФАЙЛЫ или каталоги были обнаружены. Смотрите журнал"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='green'>OK</font>"
		}
		1
		{
			$exit_code = "1"
			$exit_reason = "Один или несколько файлов были скопированы"
			$IncludeAdmin = $False
			$backupState = "<font size='3' color='green'>OK</font>"
		}
		0
		{
			$exit_code = "0"
			$exit_reason = "Файлы не скопированы, сбоев не обнаружено"
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

# Получим путь до файла, из exe файла не работает $PSScriptRoot
$path_to_script = 
	if (-not $PSScriptRoot) 
	{  
		Split-Path -Parent (Convert-Path ([environment]::GetCommandLineArgs()[0])) 
	} 
	else 
	{
		$PSScriptRoot 
	}

# Получаем данные из JSON
$json = Get-Content ( $path_to_script + '\ps_backup.json')  
$json = $json -replace '(?m)(?<=^([^"]|"[^"]*")*)//.*' -replace '(?ms)/\*.*?\*/' 
$json = $json | Out-String | ConvertFrom-Json



$json.name_dir   

$splitter = "`n_____________________________________________________________________________________________`n`n"




$path_to_backup_dir = $json.path_to_logs

# Устанавливаем и проверяем папку с логами
$path_to_logs = $path_to_backup_dir + 'logs/'
checkDir($path_to_logs)

# Устанавливаем и проверяем папку с архивами
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

#Задаем параметры для robocopy

#Для копирования папок	
$name_dir = $json.name_dir
$source = $json.source
$destination = $json.destination

#Для архивирования (заполнять необходимо из верхних строк)
# Сколько дней хранить архивы
$day_save_arc = $json.day_save_arc

#Наименование архива
$name_arc_rus = $json.name_arc_rus
	
$list_arc =  @(	#Список архвов, берется из массива робокопи
	($destination[0])
	#,($destination[1])
)

$destination_arc = $json.destination_arc 

#Перфикс к имени архива
$name_arc = $json.name_arc  

#Заполняем параметры для робокопи	
$robocopyOptions = @(
'/Z' #Копирование файлов в режиме перезапуска
,'/MIR' #Зеркально отражает дереве папок
,'/R:5' #Указывает число повторных попыток для неудавшихся копий
,'/W:15' #Указывает время ожидания между повторными попытками в секундах
,'/V' #Создает подробные выходные данные и отображает все пропущенные файлы
,'/TS' #Включает в себя отметки времени исходного файла в выходных данных 
#,'/FP' #Включает полный путь имена файлов в выходных данных
#,'/NP' #Указывает, что ход выполнения операции (число файлов или каталогов, копируются в данный момент) не будут отображаться
,'/TEE' #Записывает выходные данные о состоянии, в окне консоли, а также в файл журнала
,'/NDL' #Без списка папок
,'/NFL' #Указывает, что имена файлов не для занесения в журнал
#,'/NJH' #Указывает, что отсутствует заголовок задания.
);
[array]$strLog = $strBegin;	
[string]$date = get-date -Format "dd.MM.yyyy HH:mm:ss";
$strLog = $splitter + "**********START "+$date+"**********"+$splitter #| Out-File $fileDayLog;
[string]$sendMailText = "<center><h2> Отчет по копиям за "+$date + "</h1></center>"
$sendMailText += "<center><table border='1'   cellspacing='0' cellpadding = '5' >
  <tr >
	<td>Наименование</td>
    <td>Путь из</td>
    <td>Путь в</td>
    <td>Дата начала</td>
    <td>Сообщение</td>
    <td>Статус</td>
    <td>Дата окончания</td>
   </tr>"
for ($i = 0; $i -lt $destination.length; $i++)
{ 
	[string]$date = get-date -Format "dd.MM.yyyy HH:mm:ss";
	$sendMailText += "<tr><td>"+$name_dir[$i]+"</td><td>"+$source[$i] + "</td><td>"+$destination[$i]+"</td><td>"+$date+"</td>"
	$CmdLine = @($source[$i], $destination[$i], $fileList) + $robocopyOptions;

	#Проверим папку на наличие
	checkDir($destination[$i])
	
	#Запускаем RoboCopy с заданными параметрами
	Try	{
		$strLog +=&'robocopy.exe' $CmdLine
	}
	Catch	{
		$strLog +="Ошибка при запуске RoboCopy: "+$error[0].Exception
	}

	$tempVal = roboError
	$sendMailText += $tempVal;
	"Копирование завершено "+$name_dir[$i]+" - "+$tempVal
}
$sendMailText += "</table></center>"

#Обработка архивов
#$sendMailText +="<br><br><br><center><h2> Отчет по архивам </h2></center>"
$sendMailText += "<center><table border='1'   cellspacing='0' cellpadding = '5' >
  <tr >
	<td>Наименование</td>
    <td>Архивирование из</td>
    <td>Архивирование в</td>
    <td>Сообщение</td>
    <td>Статус</td>
    <td>Наименование</td>
	<td>Сообщение</td>
	<td>Статус</td>
   </tr>"
for ($i = 0; $i -lt $list_arc.length; $i++)
{
	checkDir($destination_arc[$i]);
	[string]$date = get-date -Format "yyyy-MM-dd";
	$fullDirArc = $destination_arc[$i]+$name_arc[$i]+$date+".zip"

	#заполняем столбцы: наименование, архивирование из, архивирование в
	$sendMailText += "<tr><td>"+$name_arc_rus[$i]+"</td><td>"+$list_arc[$i]+"</td><td>"+$fullDirArc+"</td>"
	Try{
		Compress-Archive -Path $list_arc[$i] -DestinationPath $fullDirArc -CompressionLevel Optimal -Force
		$tempVal ="Архив успешно создан "+$fullDirArc;
		$strLog +=$tempVal
		$arcStatus = "<td><font size='3' color='green'>OK</font></td>"
		}
	Catch{
		$tempVal ="Ошибка при Архивировании: "+$fullDirArc+ $error[0].Exception
		$strLog += $tempVal
		$arcStatus = "<td><font size='3' color='red'>ERROR</font></td>"
		}
	$sendMailText += "<td>"+$tempVal+"</td>"+$arcStatus
	"Архивирование завершено "+$name_arc_rus[$i]+" - "+$tempVal
	
	#Удаляем архивы старее заданной даты
	$sendMailText += "<td>Удаление архивов</td>"
	#"Проверка архивов " + $dateDelArc 
	
	#Получаем текущую дату - количество дней указанных для хранения архивов
	$dateDelArc = (get-date).addDays($day_save_arc[$i])
	#Получаем список архивов, где дата записи файла меньше даты указанной для хранения архивов
	$list_old_arc = Get-Item ($destination_arc[$i]+"*.zip") | where lastWriteTime -lt $dateDelArc
	
	#Проверяем архив на существование
	if ($list_old_arc -ne $NULL)
	{
		Try{
			#$list_old_arc | Remove-Item
			$tempVal = "Удалены архивы: "+$list_old_arc
			$strLog += $tempVal;
			$arcStatus = "<td><font size='3' color='green'>OK</font></td>"
			}
		Catch{
			$tempVal ="Ошибка при удалении архивов: "+$error[0].Exception
			$strLog +=$tempVal
			$arcStatus = "<td><font size='3' color='red'>ERROR</font></td>"
			}
	}
	else{
		$tempVal ="Нет архивов для удаления: "+$error[0].Exception
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

#$PSScriptRoot #Путь до скрипта


#Отправка почты 
#Входящие данные сообщения:
# $From = "fgdg@fgfg.com"
# $To = "fgdg@fgfg.com"
# $SMTPServer = "smtp.gmail.com"
# $SMTPPort = "587"
# $Username = "fgdg@fgfg.com"
# $Password = "pass"
# $subject = "Отчет по резервному копированию"
# $body = $sendMailText;

# #формируем сообщение в формате html:
# $message = New-Object System.Net.Mail.MailMessage $From, $To
# $message.Subject = $subject
# $message.IsBodyHTML = $true
# $message.Body = $body

# #Отправляем:

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




