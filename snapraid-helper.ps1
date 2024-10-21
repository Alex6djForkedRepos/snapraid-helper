<#
.SYNOPSIS
	SnapRAID Helper: PowerShell wrapper script for SnapRAID.
.DESCRIPTION
	This script helps automate routine SnapRAID tasks with Windows Task 
	Scheduler, ensuring that parity regularly stays in sync with your data.
.PARAMETER Argument1
	(Optional) The command to perform. Default is 'sync'.
.PARAMETER ScrubPercent
	(Optional) The percentage of data to scrub. Default is 8.33%.
.EXAMPLE
	.\snapraid-helper.ps1
	Run a regular snapraid.exe sync.
.EXAMPLE
	.\snapraid-helper.ps1 syncandscrub
	Run a sync and then a scrub.
.NOTES
	Authors: droolio, therealjmc, lrissman
	Version: 3.5-dev
	Date: 2024-10-09
#>

param(
	[string]$Argument1 = 'sync',
	[int]$ScrubPercent = 999
)

$Argument1 = $Argument1.ToLower()

$Scriptname = $MyInvocation.MyCommand.Name
#$Scriptrunning		= get-wmiobject win32_process -filter "name='powershell.exe'AND CommandLine LIKE '%$Scriptname%'"
$Scriptrunning = Get-WmiObject win32_process -Filter "name='powershell.exe'AND CommandLine LIKE '%$Scriptname%' AND NOT Handle LIKE '$PID'"
$Snapraidrunning = Get-WmiObject win32_process -Filter "name='snapraid.exe'"

$global:PreProcessHasRun = 0
$global:ServicesStarted = 0
$global:ServicesStopped = 0
$global:Diffchanges = 99
$SomethingDone = 0
$HomePath = $MyInvocation.Line | Split-Path
# General date/time format; short date, long time i.e. 'dd/MM/yyyy HH:mm:ss' but in system locale
$DateFormat = "G"
$message = ""
$ConfigError = 0

function Test-IsAdmin {
	# Borrowed from with some modifications: http://stackoverflow.com/questions/9999963/powershell-test-admin-rights-within-powershell-script
	try {
		$identity = [Security.Principal.WindowsIdentity]::GetCurrent()
		$principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $identity
		return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
	} catch {
		#throw "Failed to determine if the current user has elevated privileges. The error was: '{0}'." -f $_
		return 0
	}

	return 1
}

function Get-CurrentDate {
	$CurrentDate = Get-Date -Format $DateFormat
	return $CurrentDate.ToString()
}

function Invoke-PreRun {
	# If Process Management is enabled, then start Pre Process
	if ($config["ProcessEnable"] -eq 1 -and
		$global:PreProcessHasRun -eq 0)
	{
		# timestamp the job
		$message = "Starting Pre-Process $(Get-CurrentDate)"
		WriteLogFile $message
		$exe = $config["ProcessPre"]
		& "$exe" | Out-Null

		if (!($LastExitCode -eq "0")) {
			Invoke-PostRun
			$message = "ERROR: Pre-Process failed on $(Get-CurrentDate) with exit code $LastExitCode"
			WriteLogFile $message
			$subject = $config["SubjectPrefix"] + " " + $message
			Send-Email $subject "error" $EmailBody
			Stop-Transcript | Out-Null
			exit 1
		} else {
			$message = "Done Starting Pre-Process $(Get-CurrentDate)"
			WriteLogFile $message
			$global:PreProcessHasRun = 1
		}
	}

	Test-ParityFiles

	# If enabled take services offline
	ServiceManagement "stop"
}

function Invoke-PostRun {
	# If enabled bring services back online
	ServiceManagement "start"

	# If Process Management is enabled, then start Post Process
	if ($config["ProcessEnable"] -eq 1 -and
		$global:PreProcessHasRun -eq 1)
	{
		# timestamp the job
		$message = "Starting Post-Process $(Get-CurrentDate)"
		WriteLogFile $message
		$exe = $config["ProcessPost"]
		& "$exe" | Out-Null

		if (!($LastExitCode -eq "0")) {
			$message = "ERROR: Post-Process failed on $(Get-CurrentDate) with exit code $LastExitCode"
			WriteLogFile $message
			$subject = $config["SubjectPrefix"] + " " + $message
			Send-Email $subject "error" $EmailBody
			Stop-Transcript | Out-Null
			exit 1
		} else {
			$message = "Done Starting Post-Process $(Get-CurrentDate)"
			WriteLogFile $message
		}
	}
}

# Build Email Function (used many times in script)
function Send-Email ($fSubject, $fSuccess, $EmailBody) {
	# $fSubject -- passed subject line
	# $fSuccess -- "success" = success email, "error" = error email, "error2" = error email script/snapraid running
	$Body = ""

	$EventlogID = switch ($fSuccess) {
		"success" { 4711 }
		"error"   { 4712 }
		"error2"  { 4712 }
	}

	if ($config["IncludeExtendedInfoZip"] -eq 1 -and $fSuccess -ne "error2") {
		if (Test-Path $EmailBodyTmp) {
			Rename-Item "$EmailBodyTmp" "$EmailBodyTxt"
		}

		if (Test-Path "$EmailBodyTxt") {
			$file = Get-Item "$EmailBodyTxt"

			if ($file.length -ge $config["LogFileMaxSizeZIP"]) {
				Compress-Archive -Path "$EmailBodyTxt" -DestinationPath "$EmailBodyZip" -CompressionLevel Optimal -Force
			} else {
				$EmailBodyZip = $EmailBodyTxt
			}
		}
	}

	if ($config["EmailEnable"] -eq 1) {
		if (($fSuccess -eq "success" -and $config["EmailOnSuccess"] -eq 1) -or
			($fSuccess -eq "error" -and $config["EmailOnError"] -eq 1))
		{
			if ($config["IncludeExtendedInfo"] -eq 1) {
				$Body = (Get-Content $EmailBody | Out-String)
			}

			if ($config["IncludeExtendedInfoZip"] -eq 1) {
				if (Test-Path $EmailBodyZip) {
					if ((Get-Item $EmailBodyZip).length -le $config["MaxAttachSize"]) {
						if (Test-Path $EmailBodyZip) {
							$att = New-Object Net.Mail.Attachment ($EmailBodyZip)
							$MailMessage.Attachments.Add($att)
						}
					} else {
						Add-Content $EmailBody "LOG FILE TOO LARGE TO ATTACH"
					}
				}

				$Body = (Get-Content $EmailBody | Out-String)
			}

			$MailMessage.Subject = $fSubject
			$Mailmessage.Body = $Body
			$smtpclient.Send($MailMessage)

			if ($config["IncludeExtendedInfoZip"] -eq 1 -and (Test-Path $EmailBodyZip)) {
				$att.Dispose()
			}

		} elseif ($fSuccess -eq "error2" -and $config["EmailOnError"] -eq 1) {
			$MailMessage.Subject = $fSubject
			$Mailmessage.Body = $fSubject
			$smtpclient.Send($MailMessage)
		}
	}

	if (!(Get-EventLog -Source SnapRaid-Helper -LogName Application -ErrorAction SilentlyContinue)) {
		New-EventLog -Source SnapRaid-Helper -LogName Application
	}

	Write-EventLog -LogName Application -Source SnapRaid-Helper -EventId $EventlogID -Message $fSubject
}

function Test-ContentFiles {
	foreach ($element in $config["SnapRAIDContentFiles"]) {
		if (!(Test-Path $element)) {
			Invoke-PostRun
			$message = "ERROR: Content file ($element) not found!"
			Write-Host $message -ForegroundColor red -BackgroundColor yellow
			Add-Content $EmailBody $message
			$subject = $config["SubjectPrefix"] + " " + $message
			Send-Email $subject "error" $EmailBody
			Stop-Transcript | Out-Null
			exit 1
		}
	}
}

function Test-ParityFiles {
	foreach ($element in $config["SnapRAIDParityFiles"]) {
		if (!(Test-Path $element)) {
			Invoke-PostRun
			$message = "ERROR: Parity file ($element) not found!"
			Write-Host $message -ForegroundColor red -BackgroundColor yellow
			Add-Content $EmailBody $message
			$subject = $config["SubjectPrefix"] + " " + $message
			Send-Email $subject "error" $EmailBody
			Stop-Transcript | Out-Null
			exit 1
		}
	}
}

function WriteLogFile ($ftext) {
	Write-Host "----------------------------------------"
	Write-Host $ftext
	Write-Host "----------------------------------------"
	Add-Content $EmailBody "----------------------------------------"
	Add-Content $EmailBody $ftext
	Add-Content $EmailBody "----------------------------------------"
}

function WriteExtendedLogFile ($ftext) {
	Write-Host "----------------------------------------"
	Write-Host $ftext
	Write-Host "----------------------------------------"
	Add-Content $EmailBody "----------------------------------------"
	Add-Content $EmailBody $ftext
	Add-Content $EmailBody "----------------------------------------"

	if ($config["IncludeExtendedInfoZip"] -eq 1) {
		Add-Content $EmailBodyTmp "----------------------------------------"
		Add-Content $EmailBodyTmp $ftext
		Add-Content $EmailBodyTmp "----------------------------------------"
	}
}

function ServiceManagement ($startstop) {
	if ($startstop -eq "stop") {
		# If Service Management is enabled, then take services offline
		if ($config["ServiceEnable"] -eq 1 -and $global:ServicesStopped -ne 1) {
			# timestamp the job
			$message = "Stopping Services $(Get-CurrentDate)"
			WriteLogFile $message

			foreach ($service in $ServiceList) {
				$message = Stop-Service $service
				WriteLogFile $message
			}

			# timestamp the job
			$message = "Done Stopping Services $(Get-CurrentDate)"
			WriteLogFile $message
			$global:ServicesStopped = 1
		}
	}

	if ($startstop -eq "start") {
		# If Service Management is enabled, then bring services back online
		if ($config["ServiceEnable"] -eq 1 -and $global:ServicesStarted -ne 1 -and $global:ServicesStopped -eq 1) {
			# timestamp the job
			$message = "Starting Services $(Get-CurrentDate)"
			WriteLogFile $message

			foreach ($service in $ServiceList) {
				$message = Start-Service $service
				WriteLogFile $message
			}

			# timestamp the job
			$message = "Done Starting Services $(Get-CurrentDate)"
			WriteLogFile $message
			$global:ServicesStarted = 1
		}
	}
}

function RunSnapraid ($sargument) {
	$exe = $config["SnapRAIDPath"] + $config["SnapRAIDExe"]
	$configfile = $config["SnapRAIDPath"] + $config["SnapRAIDConfig"]

	if ($sargument -ne "fullscrub") {
		if (($ScrubPercent -ne 999) -and ($sargument -eq "scrub")) {
			& "$exe" -c $configfile $sargument -p $ScrubPercent -l $SnapRAIDLogfile 2>&1 3>&1 4>&1 | ForEach-Object { "$_" } | Tee-Object -File $TmpOutput
		} else {
			& "$exe" -c $configfile $sargument -l $SnapRAIDLogfile 2>&1 3>&1 4>&1 | ForEach-Object { "$_" } | Tee-Object -File $TmpOutput
		}
	} else {
		$sargument = "scrub"
		& "$exe" -c $configfile $sargument -p 100 -o 0 -l $SnapRAIDLogfile 2>&1 3>&1 4>&1 | ForEach-Object { "$_" } | Tee-Object -File $TmpOutput
	}

	if ($config["IncludeExtendedInfoZip"] -eq 1) {
		$FileToAdd = $EmailBodyTmp
	} else {
		$FileToAdd = $EmailBody
	}

	if (($config["ShortenLogFile"] -eq 1) -and ($sargument -ne "status")) {
		$TmpOutputInRAM = Get-Content $TmpOutput -ReadCount 0

		for ($i = 0; $i -lt $TmpOutputInRAM.length; $i++) {
			if ($TmpOutputInRAM[$i] -match "[0-9]*[A-Z]") {
				if ($TmpOutputInRAM[$i + 1] -match "[0-9]*[A-Z]") {
					$TmpOutputInRAM_First_Three = $TmpOutputInRAM[$i + 1].substring(0, 3)

					if ($TmpOutputInRAM_First_Three.substring(0, 1) -match "[0-9]") {
						if ($TmpOutputInRAM[$i].startswith($TmpOutputInRAM_First_Three)) {
						} else {
							Add-Content $FileToAdd $TmpOutputInRAM[$i]
						}
					} else {
						Add-Content $FileToAdd $TmpOutputInRAM[$i]
					}
				} else {
					if ($TmpOutputInRAM[$i + 2] -notmatch "Autosaving...") {
						Add-Content $FileToAdd $TmpOutputInRAM[$i]
					}
				}
			}
		}
	} else {
		#$TmpOutputInRAM = Get-Content $TmpOutput  -readcount 100 -delim "`0"
		# NOTE the above Get-Content command is VERY VERY VERY VERY slow, so I am using the .Net function below to get the output of the Snapraid command into a variable
		# NOTE the .Net function breaks german Umlauts so I'm using this fast way with get-content and out-string - no real time difference to .Net function
		$TmpOutputInRAM = (Get-Content $TmpOutput | Out-String)

		foreach ($line in $TmpOutputInRAM) {
			Add-Content $FileToAdd $line
			# since output is done with tee it isn't necessary to use write-host again
			# Write-Host $line
		}
	}

	if ($LastExitCode -ne "0" -and
		!($LastExitCode -eq "2" -and $sargument -eq "diff"))
	{
		Invoke-PostRun
		$message = "ERROR: SnapRAID $sargument Job FAILED on $(Get-CurrentDate) with exit code $LastExitCode"
		WriteExtendedLogFile $message
		$message2 = "Including detailed SnapRAID Log"
		WriteExtendedLogFile $message2
		$SnapRAIDLogfileInRAM = (Get-Content $SnapRAIDLogfile | Out-String)

		if ($config["IncludeExtendedInfoZip"] -eq 1) {
			$FileToAdd = $EmailBodyTmp
		} else {
			$FileToAdd = $EmailBody
		}

		foreach ($line in $SnapRAIDLogfileInRAM) {
			Add-Content $FileToAdd $line
			Write-Host $line
		}

		$subject = $config["SubjectPrefix"] + " " + $message
		Send-Email $subject "error" $EmailBody
		Stop-Transcript | Out-Null
		exit 1
	}

	# Job was successful, move onto processing.
	$message = "SnapRAID $sargument Job finished on $(Get-CurrentDate)"
	WriteExtendedLogFile $message

	if ($sargument -eq "diff") {
		DiffAnalyze
	}
}

function DiffAnalyze {
	if ($global:Diffchanges -eq 99) {
		$DEL_COUNT = Select-String $TMPOUTPUT -Pattern "^remove" | Measure-Object -Line
		$ADD_COUNT = Select-String $TMPOUTPUT -Pattern "^add" | Measure-Object -Line
		$MOVE_COUNT = Select-String $TMPOUTPUT -Pattern "^move" | Measure-Object -Line
		$RESIZE_COUNT = Select-String $TMPOUTPUT -Pattern "^resize" | Measure-Object -Line
		$UPDATE_COUNT = Select-String $TMPOUTPUT -Pattern "^update" | Measure-Object -Line

		$DEL_COUNT = $DEL_COUNT.Lines
		$ADD_COUNT = $ADD_COUNT.Lines
		$MOVE_COUNT = $MOVE_COUNT.Lines
		$UPDATE_COUNT = $UPDATE_COUNT.Lines + $RESIZE_COUNT.Lines

		$message = "SUMMARY of changes - Added [$ADD_COUNT] - Deleted [$DEL_COUNT] - Moved [$MOVE_COUNT] - Updated [$UPDATE_COUNT]"
		WriteExtendedLogFile $message

		# check if files have changed
		if ($DEL_COUNT -gt 0 -or $ADD_COUNT -gt 0 -or $MOVE_COUNT -gt 0 -or $UPDATE_COUNT -gt 0) {
			# YES, check if number of deleted files exceed DEL_THRESHOLD
			if ($DEL_COUNT -gt $config["SnapRAIDDelThreshold"]) {
				# YES, lets inform user and not proceed with the job just in case
				Invoke-PostRun
				$message = "WARNING: Number of deleted files ($DEL_COUNT) exceeded threshold (" + $config["SnapRAIDDelThreshold"] + "). NOT proceeding with job. Please run manually if this is not an error condition."
				Write-Host $message
				Add-Content $EmailBody $message
				$subject = $config["SubjectPrefix"] + " " + $message
				Send-Email $subject "error" $EmailBody
				Stop-Transcript | Out-Null
				exit 1
			} else {
				# NO, delete threshold not reached, lets run the job
				$message = "Deleted files ($DEL_COUNT) did not exceed threshold (" + $config["SnapRAIDDelThreshold"] + "), proceeding with job."
				Write-Host $message
				Add-Content $EmailBody $message
				$message = "$(Get-CurrentDate) Changes detected [A-$ADD_COUNT,D-$DEL_COUNT,M-$MOVE_COUNT,U-$UPDATE_COUNT] and deleted files ($DEL_COUNT) is below threshold (" + $config["SnapRAIDDelThreshold"] + "). Running Command."
				Write-Host $message
				Add-Content $EmailBody $message
				$global:Diffchanges = 1
			}
		} else {
			# NO, so lets log it and exit
			$message = "$(Get-CurrentDate) No change detected. Nothing to do"
			WriteExtendedLogFile $message
			$global:Diffchanges = 0
		}
	}
}

# Get variables from <scriptname>.ini
$Scriptname2 = [System.IO.Path]::GetFileNameWithoutExtension("$Scriptname")
$ConfigFile = "$HomePath\$Scriptname2.ini"
$config = @{}

Get-Content $ConfigFile | ForEach-Object {
	if (($_.startswith(";")) -or (!($_))) {
		# Non-variable or is Space
		#    write-host "Non-Variable: $_"
	} else {
		$line = $_.Split("=")
		#$config.($line[0]) = $line[1]
		$config[$line[0]] = $line[1].TrimEnd()
		#   write-host Variable: $line[0]  Content: $line[1]
	}
}

# Validate configuration variables are sane

# SnapRAID and LogFile Config
$SnapRAIDConfigs = "SnapRAIDDelThreshold", "SnapRAIDPath", "SnapRAIDExe", "SnapRAIDContentFiles", "SnapRAIDParityFiles", "TmpOutputFile", "LogFileName", "LogFileMaxSize", "LogFileZipCount", "UTF8Console", "SnapRAIDStatusAfterScrub"

foreach ($element in $SnapRAIDConfigs) {
	if (!($config[$element]) -or ($config[$element] -eq "")) {
		Write-Host "$element is null, please add a value"
		$ConfigError++
	}
}

if ($config["UTF8Console"] -eq 1) {
	chcp 65001
}

# Validate EmailBodyPath and if not specified, use ScriptPath
if (!($config["LogPath"]) -or ($config["LogPath"] -eq "")) {
	$config["LogPath"] = "$HomePath\"
}

if (!(Test-Path $config["LogPath"] -PathType container)) {
	Write-Host "ERROR: LogPath: " $config["LogPath"] "  - Path Does not exist.  Please fix $ConfigFile or create the path"
	exit 1
} else {
	$LogPathTest = $config["LogPath"].EndsWith("\")

	if (!($LogPathTest)) {
		$config["LogPath"] = $config["LogPath"] += "\"
	}
}

$LogFile = $config["LogPath"] + $config["LogFileName"]

# Email Configs
$EmailConfigs = "SubjectPrefix", "EmailTo", "EmailFrom", "Body", "SMTPHost", "SMTPSSLEnable", "SMTPAuthEnable", "EmailBodyFile", "EmailBodyFileZip", "EmailEnable", "SMTPPort", "EmailOnSuccess", "EmailOnError", "IncludeExtendedInfo", "IncludeExtendedInfoZip", "LogFileMaxSizeZIP", "MaxAttachSize", "ShortenLogFile"

# If email is enabled, validate email configs are not null
if ($config["EmailEnable"] -eq 1) {
	foreach ($element in $EmailConfigs) {
		if (!($config[$element]) -or ($config[$element] -eq "")) {
			Write-Host "$element is null, please add a value"
			$ConfigError++
		}
	}

	if ($config["IncludeExtendedInfoZip"] -eq 1) {
		$config["IncludeExtendedInfo"] = 0
	}
}

# Service/Process Configs
if ($config["ServiceEnable"] -eq 1) {
	if (!(Test-IsAdmin)) {
		Write-Host "You need to run the script with elevated rights to start and stop services. Either run with Elevated Rights or change in $ConfigFile ProcessEnable=0"
		exit 1
	}

	$ServiceConfigs = "ServiceName"

	# If service handling is enabled, validate configs are not null
	foreach ($element in $ServiceConfigs) {
		if (!($config[$element]) -or ($config[$element] -eq "")) {
			Write-Host "$element is null, please add a value"
			$ConfigError++
		}
	}

	$ServiceNum = 0
	$ServiceList = $config["ServiceName"].Split(",").Replace('"', "")

	foreach ($Service in $ServiceList) {
		$ServiceNum++

		if (!(Get-Service $Service -ErrorAction SilentlyContinue)) {
			"The Service $Service does not exist.   Please remove or correct in ProcessName in $ConfigFile"
		}
	}
}

if ($config["ProcessEnable"] -eq 1) {
	# If process handling is enabled, validate configs are not null

	$ProcessConfigs = "ProcessPre", "ProcessPost"

	foreach ($element in $ProcessConfigs) {
		if (!($config[$element]) -or ($config[$element] -eq "")) {
			Write-Host "$element is null, please add a value"
			$ConfigError++
		}

		if (!(Test-Path $config[$element])) {
			wite-host "$config[$element] is not a valid path to execute!"
			$ConfigError++
		}
	}
}

# EventLog Configs
$EventLogConfigs = "EventLogSources", "EventLogEntryType", "EventLogDays", "EventLogHaltOnDiskError"
if ($config["EventLogEnable"] -eq 1) {
	foreach ($element in $EventLogConfigs) {
		if (!($config[$element]) -or ($config[$element] -eq "")) {
			Write-Host "$element is null, please add a value"
			$ConfigError++
		}
	}

	$EventLogEntryTypeList = $config["EventLogEntryType"].Replace('"', "").Trim().Split(",")
	$EventLogSourcesList = $config["EventLogSources"].Replace('"', "").Trim().Split(",")
}

# Report if there are errors and exit
if ($ConfigError -ge 1) {
	Write-Host "Number of config errors: $ConfigError"
	Write-Host "Please correct $ConfigFile and run again"
	exit 1
}

# Validate EmailBodyPath and if not specified, use Windows Temp path
if (!($config["EmailBodyPath"]) -or ($config["EmailBodyPath"] -eq "")) {
	$config["EmailBodyPath"] = "$env:temp\"
}

if (!(Test-Path $config["EmailBodyPath"] -PathType container)) {
	Write-Host "ERROR: EmailBodyPath: " $config["EmailBodyPath"] "  - Path Does not exist.  Please fix $ConfigFile or create the path"
	exit 1
}

# Validate TmpOutputPath and if not specified, use Windows Temp path
if (!($config["TmpOutputPath"]) -or ($config["TmpOutputPath"] -eq "")) {
	$config["TmpOutputPath"] = "$env:temp\"
}

if (!(Test-Path $config["TmpOutputPath"] -PathType container)) {
	Write-Host "ERROR: TmpOutputPath:" $config["TmpOutputPath"] "  - Path Does not exist.  Please fix $ConfigFile or create the path"
	exit 1
}

# Initalize Email
if ($config["EmailEnable"] -eq 1) {
	$SMTPClient = New-Object Net.Mail.SmtpClient
	$MailMessage = New-Object Net.Mail.Mailmessage
	$MailMessage.IsBodyHtml = $false
	$SMTPClient.Host = $config["SMTPHost"]
	$SMTPClient.Port = $config["SMTPPort"]
	$MailMessage.From = $config["EmailFrom"]
	$MailMessage.To.Add($config["EmailTo"])

	if ($config["SMTPSSLEnable"] -eq 1) {
		$SMTPClient.EnableSsl = $true
		[System.Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
	} else {
		$SMTPClient.EnableSsl = $false
	}

	if ($config["SMTPAuthEnable"] -eq 1) {
		$SMTPClient.Credentials = New-Object System.Net.NetworkCredential ($config["SMTPUID"], $config["SMTPPass"]);
	}
}

$TmpOutput = $config["TmpOutputPath"] + $config["TmpOutputfile"]
$EmailBody = $config["EmailBodyPath"] + $config["EmailBodyfile"]
$EmailBodyTmp = $config["EmailBodyPath"] + $config["EmailBodyFileZip"] + ".out"
$EmailBodyTxt = $config["EmailBodyPath"] + $config["EmailBodyFileZip"] + ".txt"
$EmailBodyZip = $config["EmailBodyPath"] + $config["EmailBodyFileZip"] + ".zip"
$SnapRAIDLogfile = $config["TmpOutputPath"] + "snapRAIDerror.out"

# Ensure only one Snapraid process and only one instance of this script is running
# Note that the detection for running script only works if it is called with the script as a parameter
# for example powershell.exe snapraid-helper.ps1 - but not if the script is called like .\snapraid-helper.ps1
if ($Scriptrunning -match "Handle") {
	$message = "ERROR: Another instance of the script is still running! $Argument1 can't run on $(Get-CurrentDate)"
	Write-Host "----------------------------------------"
	Write-Host $message
	Write-Host "----------------------------------------"
	$subject = $config["SubjectPrefix"] + " " + $message
	Send-Email $subject "error2"
	exit 1
}

if ($Snapraidrunning -match "Handle") {
	$message = "ERROR: Another instance of snapraid is still running! $Argument1 can't run on $(Get-CurrentDate)"
	Write-Host "----------------------------------------"
	Write-Host $message
	Write-Host "----------------------------------------"
	$subject = $config["SubjectPrefix"] + " " + $message
	Send-Email $subject "error2"
	exit 1
}

# Start with some cleanup
if (Test-Path $TmpOutput) {
	Remove-Item $TmpOutput
}

if (Test-Path $EmailBody) {
	Remove-Item $EmailBody
}

if (Test-Path $EmailBodyTmp) {
	Remove-Item $EmailBodyTmp
}

if (Test-Path $EmailBodyTxt) {
	Remove-Item $EmailBodyTxt
}

if (Test-Path $EmailBodyZip) {
	Remove-Item $EmailBodyZip
}

if (Test-Path $SnapRAIDLogfile) {
	Remove-Item $SnapRAIDLogfile
}

if ($config["EnableDebugOutput"] -eq 1) {
	foreach ($element in $Config) {
		Write-Output $element
		Write-Output "TmpOutput = $TmpOutput"
		Write-Output "EmailBody = $EmailBody"
		Write-Output "EmailBodyTmp = $EmailBodyTmp"
		Write-Output "EmailBodyTxt = $EmailBodyTxt"
		Write-Output "EmailBodyZip = $EmailBodyZip"
		Write-Output "SnapRAIDLogfile = $SnapRAIDLogfile"
	}
}

# Log Management Section
if (Test-Path "$LogFile") {
	$file = Get-Item "$LogFile"

	if ($file.length -ge $config["LogFileMaxSize"]) {
		if ($config["LogFileZipCount"] -ge 1) {
			$i = $config["LogFileZipCount"]

			if (Test-Path "$LogFile.$i.zip") {
				Remove-Item "$LogFile.$i.zip"
			}

			while ($i -gt 1) {
				$j = $i - 1

				if (Test-Path "$LogFile.$j.zip") {
					Rename-Item "$LogFile.$j.zip" "$LogFile.$i.zip"
				}

				$i = $i - 1
			}

			Compress-Archive -Path "$LogFile" -DestinationPath "$LogFile.zip" -CompressionLevel Optimal -Force
			Rename-Item "$LogFile.zip" "$LogFile.1.zip"
		}

		Remove-Item "$LogFile"
		New-Item "$LogFile" -Type file
	}
}

# Start Transcript logging
# redirect all stdout to log file (leave stderr alone thou)
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -Path $LogFile -Append

# Check Eventlog for Errors
$message = "Checking for Disk issues in Eventlog at $(Get-CurrentDate)"
WriteLogFile $message

$EventLogOutput = Get-EventLog -LogName system -EntryType $EventLogEntryTypeList -Source $EventLogSourcesList -After (Get-Date).AddDays($config["EventLogdays"])
Write-Host "TimeGenerated,EntryType,Source,Message"

foreach ($event in $EventLogOutput) {
	$EventLogCount = $EventLogcount + 1
	$TimeGenerated = $event.TimeGenerated
	$EntryType = $event.EntryType
	$Source = $event.Source
	$EventMessage = $event.Message

	Write-Host "$TimeGenerated,$EntryType,$Source,$EventMessage"
	Add-Content $EmailBody "$TimeGenerated,$EntryType,$Source,$EventMessage"
}

if (($EventLogCount -ge 1) -and ($config["EventLogHaltOnDiskError"] -eq 1)) {
	$message = "WARN: Found disk Errors/Warnings in EventLogs.  Aborting sync based on HaltOnDiskError"
	Write-Host $message -ForegroundColor red -BackgroundColor yellow
	Add-Content $EmailBody $message
	$subject = $config["SubjectPrefix"] + " " + $message
	Send-Email $subject "error" $EmailBody
	Stop-Transcript | Out-Null
	exit 1
}

# sanity check first to make sure we can access the content and parity files
$config["SnapRAIDContentFiles"] = $config["SnapRAIDContentFiles"].Split(",")
$config["SnapRAIDParityFiles"] = $config["SnapRAIDParityFiles"].Split(",")

Test-ContentFiles

# timestamp the job
$message = "SnapRAID $argument1 Job started on $(Get-CurrentDate)"
WriteExtendedLogFile $message

if ($Argument1 -eq "syncandcheck" -and $SomethingDone -ne 1) {
	$argument = "diff"
	RunSnapraid $argument

	Invoke-PreRun

	if ($global:Diffchanges -eq 1) {
		$argument = "sync"
		RunSnapraid $argument
	}

	$argument = "check"
	RunSnapraid $argument
	Invoke-PostRun
	$message = "SUCCESS: SnapRAID SYNC and CHECK Job finished on $(Get-CurrentDate)"
	WriteExtendedLogFile $message
	$subject = $config["SubjectPrefix"] + " " + $message
	Send-Email $subject "success" $EmailBody
	$SomethingDone = 1

} elseif ($Argument1 -eq "syncandscrub" -and $SomethingDone -ne 1) {
	$argument = "diff"
	RunSnapraid $argument

	Invoke-PreRun

	if ($global:Diffchanges -eq 1) {
		$argument = "sync"
		RunSnapraid $argument
	}

	$argument = "scrub"
	RunSnapraid $argument

	if ($config["SnapRAIDStatusAfterScrub"] -eq 1) {
		$argument = "status"
		RunSnapraid $argument
	}

	Invoke-PostRun
	$message = "SUCCESS: SnapRAID SYNC and SCRUB Job finished on $(Get-CurrentDate)"
	WriteExtendedLogFile $message
	$subject = $config["SubjectPrefix"] + " " + $message
	Send-Email $subject "success" $EmailBody
	$SomethingDone = 1

} elseif ($Argument1 -eq "syncandfix" -and $SomethingDone -ne 1) {
	$argument = "diff"
	RunSnapraid $argument

	Invoke-PreRun

	if ($global:Diffchanges -eq 1) {
		$argument = "sync"
		RunSnapraid $argument
	}

	$argument = "fix"
	RunSnapraid $argument
	Invoke-PostRun
	$message = "SUCCESS: SnapRAID SYNC and FIX Job finished on $(Get-CurrentDate)"
	WriteExtendedLogFile $message
	$subject = $config["SubjectPrefix"] + " " + $message
	Send-Email $subject "success" $EmailBody
	$SomethingDone = 1

} elseif ($Argument1 -eq "syncandfullscrub" -and $SomethingDone -ne 1) {
	$argument = "diff"
	RunSnapraid $argument

	Invoke-PreRun

	if ($global:Diffchanges -eq 1) {
		$argument = "sync"
		RunSnapraid $argument
	}

	$argument = "fullscrub"
	RunSnapraid $argument
	Invoke-PostRun
	$message = "SUCCESS: SnapRAID SYNC and FULL SCRUB Job finished on $(Get-CurrentDate)"
	WriteExtendedLogFile $message
	$subject = $config["SubjectPrefix"] + " " + $message
	Send-Email $subject "success" $EmailBody
	$SomethingDone = 1
}

if ($SomethingDone -ne 1) {
	# If another command was passed to the script run this command, else run the sync command
	if ($Argument1 -ne "sync") {
		if ($Argument1 -ne "diff" -and $Argument1 -ne "list" -and $Argument1 -ne "dup" -and $Argument1 -ne "status" -and $Argument1 -ne "pool") {
			Invoke-PreRun
		}

		$argument = $Argument1
		RunSnapraid $argument
		Invoke-PostRun
		$message = "SUCCESS: SnapRAID $Argument1 Job finished on $(Get-CurrentDate)"
		WriteExtendedLogFile $message
		$subject = $config["SubjectPrefix"] + " " + $message
		Send-Email $subject "success" $EmailBody
		$SomethingDone = 1

	} else {
		$argument = "diff"
		RunSnapraid $argument

		if ($global:Diffchanges -eq 1) {
			Invoke-PreRun
			$argument = "sync"
			RunSnapraid $argument
			Invoke-PostRun
			$message = "SUCCESS: SnapRAID SYNC Job finished on $(Get-CurrentDate)"
			WriteExtendedLogFile $message
			$subject = $config["SubjectPrefix"] + " " + $message
			Send-Email $subject "success" $EmailBody
			$SomethingDone = 1

		} else {
			# NO, so lets log it and exit
			Invoke-PostRun
			$message = "$(Get-CurrentDate) No change detected. Nothing to do"
			WriteExtendedLogFile $message
			$subject = $config["SubjectPrefix"] + " SUCCESS: SnapRAID SYNC - No change detected. Nothing to do"
			Send-Email $subject "success" $EmailBody
			$SomethingDone = 1
		}
	}
}

# End Transcript
Stop-Transcript | Out-Null
