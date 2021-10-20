<#
	.Synopsis
		Script d'alerte pour fermeture de journ�e manqu�e par Ma�tre'D.
	.Description
		Ce script v�rifie dans un r�pertoire donn� pour la pr�sence de l'archive quotidienne de fermeture g�n�r�e par Ma�tre'D. Dans le cas o� l'archive est manquante, une alerte est pouss�e pour indiquer l'archive manquante.
	.Parameter Path
		Chemin d'acc�s menant au r�pertoire o� Ma�tre'D d�pose les archives g�n�r�es � la fermeture d'une journ�e.
		Mandatoire, doit �tre valide pour l'utilisateur ex�cutant le script.
	.Parameter ClosingDate
		Date de fermeture recherch�e.
		Non mandatoire, valeur par d�faut � la date courante.
	.Parameter SMTPServer
		Adresse du serveur de courriels � utiliser pour envoyer l'alerte courriel.
		Mandatoire.
	.Parameter SMTPPort
		Num�ro de port du serveur de courriels � utiliser pour envoyer l'alerte courriel.
		Non mandatoire, valeur par d�faut � 25.
	.Parameter EmailTo
		Adresse courriel � qui envoyer l'alerte. G�n�ralement une liste de distribution o� les membres doivent �tre au courant du probl�me potentiel.
		Mandatoire, doit �tre un format d'adresse courriel reconnu par .NET.
	.Parameter EmailFrom
		Adresse courriel � partir de qui envoyer l'alerte. G�n�ralement le compte courriel r�serv� aux alertes.
		Mandatoire, doit �tre un format d'adresse courriel reconnu par .NET.
	.Parameter EmailSubject
		Titre � utiliser pour le message �lectronique d'alerte.
		Non mandatoire, titre g�n�ral informatif avec la date manqu�e par Ma�tre�D.
	.Parameter EmailBody
		Corps du message � utiliser pour le courriel d'alerte.
		Non mandatoire, contenu g�n�ral informatif avec informations pertinentes par d�faut.
	.Parameter HealthcheckURL
		URL � pinger pour indiquer que la t�che planifi�e est fonctionnelle.
		Non mandatoire.
	.Inputs
		Aucun. Vous ne pouvez pas passer d'objets � ce script.
	.Outputs
		Aucun. Seulement des messages informatifs sont affich�s � la console.
	.Example
		PS> .\Check-MissingDayClosureFromMaitreD.ps1 -Path . -SMTPServer smtp.local -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local
	.Example
		ST> powershell -NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File ".\Check-MissingDayClosureFromMaitreD.ps1" -Path . -SMTPServer smtp.local -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local
	.Example
		PS> .\Check-MissingDayClosureFromMaitreD.ps1 -Path . -ClosingDate 2021-09-23 -SMTPServer smtp.local -SMTPPort 25 -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local -EmailSubject "Fermeture manqu�e par Ma�tre'D" -EmailBody "Une fermeture semble avoir �t� manqu�e par Ma�tre'D. S.v.p. v�rifier."
#>

# Raccourcis de formats d'affichage de dates : https://www.tutorialspoint.com/how-to-format-date-string-in-powershell
# Localisation des affichages de dates : https://community.idera.com/database-tools/powershell/powertips/b/tips/posts/formatting-date-and-time-with-culture

[CmdletBinding()]
Param (
	# R�pertoire � surveiller
	[Parameter(Mandatory = $true)]
	[ValidateScript({ Test-Path $_ -PathType Container })]
	[System.IO.DirectoryInfo] $Path,
	
	# Date � v�rifier
	[Parameter(Mandatory = $false)]
	[ValidateScript({ [bool](([datetime] $_).GetType().name -match "DateTime") })]
	[string] $ClosingDate = $(Get-Date),
	
	# Serveur SMTP � utiliser pour envoyer le courriel d'alerte
	[Parameter(Mandatory = $true)]
	[string] $SMTPServer,

	# Port SMTP � utiliser pour envoyer le courriel d'alerte
	[Parameter(Mandatory = $false)]
	[int] $SMTPPort = 25,
	
	# Adresse courriel � laquelle envoyer l'alerte
	# 
	# Validation inspir�e de la r�ponse de Mathias R. Jessen sur Stackoverflow
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
	
	# Adresse courriel � partir de laquelle envoyer l'alerte
	# 
	# Validation inspir�e de la r�ponse de Mathias R. Jessen sur Stackoverflow
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
	
	# Titre de l'alerte � envoyer
	[Parameter(Mandatory = $false)]
	[ValidateNotNullOrEmpty()]
	[string] $EmailSubject = "Fermeture manquante pour Ma�tre'D : $((Get-Date ([datetime]$script:ClosingDate)).ToString("d", [CultureInfo] "fr-CA"))",
	
	# Corps de l'alerte � envoyer
	[Parameter(Mandatory = $false)]
	[string] $EmailBody = "La journ�e du $((Get-Date ([datetime]$script:ClosingDate)).ToString("D", [CultureInfo] "fr-CA")) semble �tre manquante.`nV�rifiez que les tables sont belles et bien ferm�es puis tentez la fermeture de la journ�e.`n`nServeur d'origine : $env:ComputerName`nR�pertoire d'origine : $($script:Path.FullName)`nDate de fermeture recherch�e : $((Get-Date ([datetime]$script:ClosingDate)).ToString("d", [CultureInfo] "fr-CA"))",

	# URL de Healthcheck � pinger pour indiquer que le script a roul�
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
			Write-Host "Ping du healthcheck pour d�marrage..."
			Invoke-RestMethod "$script:HealthcheckURL/start"
		}
		elseif ($Fail)
		{
			Write-Host "Ping du healthcheck pour �chec..."
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
Write-Debug "Chemin � surveiller : $script:Path"
Write-Debug "Date de fermeture � v�rifier : $script:ClosingDateAsObject"
Write-Debug "Serveur SMTP � utiliser pour alerter : $script:SMTPServer"
Write-Debug "Port SMTP � utiliser pour alerter : $script:SMTPPort"
Write-Debug "Adresse courriel � alerter : $script:EmailTo"
Write-Debug "Adresse courriel � partir de laquelle alerter : $script:EmailFrom"
Write-Debug "Sujet du message d'alerte : $script:EmailSubject"
Write-Debug "Contenu du message d'alerte : $script:EmailBody"
Write-Debug "URL de Healthcheck : $script:HealthcheckURL"
Write-Debug "Healthcheck sera utilis� : $(if ($script:HealthcheckURL) {$true} else {$false})"
Write-Host "D�but du script."
New-ConsoleLine

Ping-Healthcheck -Start
try
{
	Write-Host "Obtention des fichiers dans le r�pertoire $($script:Path.Name) pour la fermeture du $((Get-Date $script:ClosingDateAsObject).ToString("D", [CultureInfo] "fr-CA"))..."
	[System.IO.FileInfo[]] $Archives = Get-ChildItem -Path $script:Path | Where-Object { ($script:ClosingDateAsObject.Date -eq $_.CreationTime.Date) -and ($_.Name -like "$(Get-Date -Date $script:ClosingDateAsObject -Day ($script:ClosingDateAsObject.Day - 1) -Format "yyyyMMdd")*") }
	Write-Debug "Nombre de fichiers trouv�s : $($script:Archives.Length)"
	if ($script:Archives.Length -gt 0)
	{
		# On a trouv� au moins un fichier!
		Write-Debug "Fichiers trouv�s : $script:Archives"
		Write-Host "Les fichiers semblent exister!"
		New-ConsoleLine
	}
	else
	{
		# On a rien trouv�!
		Write-Error "Aucun fichier n'a �t� trouv� dans le r�pertoire $($script:Path.FullName) pour la fermeture du $((Get-Date $script:ClosingDateAsObject).ToString("D", [CultureInfo] "fr-CA"))!"
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
