<#
	.Synopsis
		Script d'alerte pour fermeture de journée manquée par Maître'D.
	.Description
		Ce script vérifie dans un répertoire donné pour la présence de l'archive quotidienne de fermeture générée par Maître'D. Dans le cas où l'archive est manquante, une alerte est poussée pour indiquer l'archive manquante.
	.Parameter Path
		Chemin d'accès menant au répertoire où Maître'D dépose les archives générées à la fermeture d'une journée.
		Mandatoire, doit être valide pour l'utilisateur exécutant le script.
	.Parameter ClosingDate
		Date de fermeture recherchée.
		Non mandatoire, valeur par défaut à la date courante.
	.Parameter SMTPServer
		Adresse du serveur de courriels à utiliser pour envoyer l'alerte courriel.
		Mandatoire.
	.Parameter SMTPPort
		Numéro de port du serveur de courriels à utiliser pour envoyer l'alerte courriel.
		Non mandatoire, valeur par défaut à 25.
	.Parameter EmailTo
		Adresse courriel à qui envoyer l'alerte. Généralement une liste de distribution où les membres doivent être au courant du problème potentiel.
		Mandatoire, doit être un format d'adresse courriel reconnu par .NET.
	.Parameter EmailFrom
		Adresse courriel à partir de qui envoyer l'alerte. Généralement le compte courriel réservé aux alertes.
		Mandatoire, doit être un format d'adresse courriel reconnu par .NET.
	.Parameter EmailSubject
		Titre à utiliser pour le message électronique d'alerte.
		Non mandatoire, titre général informatif avec la date manquée par Maître’D.
	.Parameter EmailBody
		Corps du message à utiliser pour le courriel d'alerte.
		Non mandatoire, contenu général informatif avec informations pertinentes par défaut.
	.Parameter HealthcheckURL
		URL à pinger pour indiquer que la tâche planifiée est fonctionnelle.
		Non mandatoire.
	.Inputs
		Aucun. Vous ne pouvez pas passer d'objets à ce script.
	.Outputs
		Aucun. Seulement des messages informatifs sont affichés à la console.
	.Example
		PS> .\Check-MissingDayClosureFromMaitreD.ps1 -Path . -SMTPServer smtp.local -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local
	.Example
		ST> powershell -NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File ".\Check-MissingDayClosureFromMaitreD.ps1" -Path . -SMTPServer smtp.local -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local
	.Example
		PS> .\Check-MissingDayClosureFromMaitreD.ps1 -Path . -ClosingDate 2021-09-23 -SMTPServer smtp.local -SMTPPort 25 -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local -EmailSubject "Fermeture manquée par Maître'D" -EmailBody "Une fermeture semble avoir été manquée par Maître'D. S.v.p. vérifier."
#>

# Raccourcis de formats d'affichage de dates : https://www.tutorialspoint.com/how-to-format-date-string-in-powershell
# Localisation des affichages de dates : https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/formatting-date-and-time-with-culture

[CmdletBinding()]
Param (
	# Répertoire à surveiller
	[Parameter(Mandatory = $true)]
	[ValidateScript({ Test-Path $_ -PathType Container })]
	[System.IO.DirectoryInfo] $Path,
	
	# Date à vérifier
	[Parameter(Mandatory = $false)]
	[ValidateScript({ [bool](([datetime] $_).GetType().name -match "DateTime") })]
	[string] $ClosingDate = $(Get-Date),
	
	# Serveur SMTP à utiliser pour envoyer le courriel d'alerte
	[Parameter(Mandatory = $true)]
	[string] $SMTPServer,

	# Port SMTP à utiliser pour envoyer le courriel d'alerte
	[Parameter(Mandatory = $false)]
	[int] $SMTPPort = 25,
	
	# Adresse courriel à laquelle envoyer l'alerte
	# 
	# Validation inspirée de la réponse de Mathias R. Jessen sur Stackoverflow
	# https://stackoverflow.com/a/48254513
	[Parameter(Mandatory = $true)]
	[ValidateScript({ 
			try 
			{
				$null = [mailaddress] $_
				return $true
			}
			catch
			{
				return $false
			}
		})]
	[string] $EmailTo,
	
	# Adresse courriel à partir de laquelle envoyer l'alerte
	# 
	# Validation inspirée de la réponse de Mathias R. Jessen sur Stackoverflow
	# https://stackoverflow.com/a/48254513
	[Parameter(Mandatory = $true)]
	[ValidateScript({ 
			try 
			{
				$null = [mailaddress] $_
				return $true
			}
			catch
			{
				return $false
			}
		})]
	[string] $EmailFrom,
	
	# Titre de l'alerte à envoyer
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string] $EmailSubject = "Fermeture manquante pour Maître'D : $((Get-Date ([datetime]$script:ClosingDate)).ToString("d", [CultureInfo] "fr-CA"))",
	
	# Corps de l'alerte à envoyer
	[Parameter(Mandatory = $false)]
	[string] $EmailBody = "La journée du $((Get-Date ([datetime]$script:ClosingDate)).ToString("D", [CultureInfo] "fr-CA")) semble être manquante.`nVérifiez que les tables sont belles et bien fermées puis tentez la fermeture de la journée.`n`nServeur d'origine : $env:ComputerName`nRépertoire d'origine : $($script:Path.FullName)`nDate de fermeture recherchée : $((Get-Date ([datetime]$script:ClosingDate)).ToString("d", [CultureInfo] "fr-CA"))",

	# URL de Healthcheck à pinger pour indiquer que le script a roulé
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string] $HealthcheckURL
)

function New-ConsoleLine
{
	Write-Host
}

function Send-Alert
{
	Write-Host "Envoi de la notification d'alerte..."
	Send-MailMessage -SmtpServer $script:SMTPServer -Port $script:SMTPPort -From $script:EmailFrom -To $script:EmailTo -Subject $script:EmailSubject -Body $script:EmailBody -Priority High -Encoding UTF8
	Write-Host OK

	New-ConsoleLine
}

function Ping-Healthcheck
{
	param (
		[Switch] $Start,
		[Switch] $Fail
	)

	if ($script:HealthcheckURL)
	{
		if ($Start)
		{
			Write-Host "Ping du healthcheck pour démarrage..."
			Invoke-RestMethod "$script:HealthcheckURL/start"
		}
		elseif ($Fail)
		{
			Write-Host "Ping du healthcheck pour échec..."
			Invoke-RestMethod "$script:HealthcheckURL/fail"
		}
		else
		{
			Write-Host "Ping du healthcheck..."
			Invoke-RestMethod $script:HealthcheckURL
		}
	}

	New-ConsoleLine
}

[datetime] $ClosingDateAsObject = [datetime] $script:ClosingDate
Write-Debug "Chemin à surveiller : $script:Path"
Write-Debug "Date de fermeture à vérifier : $script:ClosingDateAsObject"
Write-Debug "Serveur SMTP à utiliser pour alerter : $script:SMTPServer"
Write-Debug "Port SMTP à utiliser pour alerter : $script:SMTPPort"
Write-Debug "Adresse courriel à alerter : $script:EmailTo"
Write-Debug "Adresse courriel à partir de laquelle alerter : $script:EmailFrom"
Write-Debug "Sujet du message d'alerte : $script:EmailSubject"
Write-Debug "Contenu du message d'alerte : $script:EmailBody"
Write-Debug "URL de Healthcheck : $script:HealthcheckURL"
Write-Debug "Healthcheck sera utilisé : $(if ($script:HealthcheckURL) {$true} else {$false})"
Write-Host "Début du script."
New-ConsoleLine

Ping-Healthcheck -Start
try
{
	Write-Host "Obtention des fichiers dans le répertoire $($script:Path.Name) pour la fermeture du $((Get-Date $script:ClosingDateAsObject).ToString("D", [CultureInfo] "fr-CA"))..."
	[System.IO.FileInfo[]] $Archives = Get-ChildItem -Path $script:Path | Where-Object { ($script:ClosingDateAsObject.Date -eq $_.CreationTime.Date) -and ($_.Name -like "$(Get-Date -Date $script:ClosingDateAsObject -Day ($script:ClosingDateAsObject.Day - 1) -Format "yyyyMMdd")*") }
	Write-Debug "Nombre de fichiers trouvés : $($script:Archives.Length)"
	if ($script:Archives.Length -gt 0)
	{
		# On a trouvé au moins un fichier!
		Write-Debug "Fichiers trouvés : $script:Archives"
		Write-Host "Les fichiers semblent exister!"
		New-ConsoleLine
	}
	else
	{
		# On a rien trouvé!
		Write-Error "Aucun fichier n'a été trouvé dans le répertoire $($script:Path.FullName) pour la fermeture du $((Get-Date $script:ClosingDateAsObject).ToString("D", [CultureInfo] "fr-CA"))!"
		Send-Alert
		New-ConsoleLine
	}

	Ping-Healthcheck
}
catch
{
	Ping-Healthcheck -Fail
}

Write-Host "Fin du script."
Exit
