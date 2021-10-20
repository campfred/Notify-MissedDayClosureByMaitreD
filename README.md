# Notify-MissedDayClosureByMaitreD

Script PowerShell pour rechercher la présence d'un fichier d'archive produit lors de la fermeture d'une journée dans Maître'D et avertir dans le cas de son absence.

### Paramètres

| Nom              | Mandatoire | Description                                                  | Valeur par défaut                                            |
| ---------------- | ---------- | ------------------------------------------------------------ | ------------------------------------------------------------ |
| `Path`           | Oui        | Chemin d'accès menant au répertoire où Maître'D dépose les archives générées à la fermeture d'une journée. |                                                              |
| `ClosingDate`    | Non        | Date de fermeture recherchée.                                | Date courante                                                |
| `SMTPServer`     | Oui        | Adresse du serveur de courriels à utiliser pour envoyer l'alerte courriel. |                                                              |
| `SMTPPort`       | Non        | Numéro de port du serveur de courriels à utiliser pour envoyer l'alerte courriel. | 25                                                           |
| `EmailTo`        | Oui        | Adresse courriel à qui envoyer l'alerte. Généralement une liste de distribution où les membres doivent être au courant du problème potentiel. |                                                              |
| `EmailFrom`      | Oui        | Adresse courriel à partir de qui envoyer l'alerte. Généralement le compte courriel réservé aux alertes. |                                                              |
| `EmailSubject`   | Non        | Titre à utiliser pour le message électronique d'alerte.      | Titre général informatif avec la date manquée par Maître’D.  |
| `EmailBody`      | Non        | Corps du message à utiliser pour le courriel d'alerte.       | Contenu général informatif avec informations pertinentes comme date recherchée, répertoire, nom d’hôte. |
| `HealthcheckURL` | Non        | URL à pinger pour indiquer que la tâche planifiée est fonctionnelle. |                                                              |

## Prérequis

| Logiciel | Version |
| --- | --- |
| PowerShell | 3+ |

## Instructions

### Manuellement ou via script

Lancer le script avec la commande `.\Notify-MissedDayClosureByMaitreD.ps1 -Path . -SMTPServer smtp.local -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local`.

> Ajuster les paramètres ainsi que le chemin du script en fonction de l’environnement utilisé.

### Par tâche planifiée

Lancer le script avec une action de lancement de programmet et utiliser le chemin `powershell -NoProfile -NoLogo -NonInteractive -ExecutionPolicy Bypass -File ".\Notify-MissedDayClosureByMaitreD.ps1" -Path . -SMTPServer smtp.local -EmailFrom alerte@smtp.local -EmailTo alertemaitred@smtp.local` pour l’exécuter.

> Ajuster les paramètres ainsi que le chemin du script en fonction de l’environnement utilisé.

