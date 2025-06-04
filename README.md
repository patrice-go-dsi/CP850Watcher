# Script PowerShell pour transcodage automatique de fichiers UTF8 en CP850 (imports SAGE)

## Description

Ce dépôt contient un script PowerShell qui surveille un dossier de sortie d’un client (ex. `c:\Work\IN`) afin de transcoder automatiquement les fichiers texte en encodage **CP850**, requis pour leur intégration dans **SAGE**.

Le script convertit les fichiers au format CP850 et les enregistre dans le répertoire `c:\Work\OUT`.

## Fonctionnalités

- Surveillance en temps réel des fichiers `*.txt` créés, modifiés ou renommés.
- Transcodage automatique de UTF-8 vers CP850.
- Écriture des fichiers transcodés dans un répertoire cible.
- Journalisation des opérations dans un fichier log.
- Fonctionne comme un service Windows via **WinSW**, avec démarrage automatique.

## Emplacement des fichiers

| Élément                       | Emplacement                                        |
|-------------------------------|----------------------------------------------------|
| Script PowerShell             | `C:\Scripts\watcher-cp850.ps1`                     |
| Répertoire surveillé          | `C:\Work\IN` (à adapter selon le besoin)           |
| Répertoire de sortie (CP850)  | `C:\Work\OUT`                                      |
| Logs                          | `C:\Scripts\Logs`                                  |
| Exécutable du service WinSW   | `C:\Services\CP850Watcher\CP850WatcherService.exe` |

## Paramétrage

Les chemins de répertoires ainsi que les filtres de fichiers sont définis dans le script PowerShell lui-même (`watcher-cp850.ps1`). 
Il est possible de les modifier directement dans ce fichier en cas de besoin.

## Gestion du service

Le script est encapsulé dans un service Windows grâce à [WinSW (Windows Service Wrapper)](https://github.com/winsw/winsw).  
Le service démarre automatiquement avec Windows.

### Commandes utiles

Si le script PowerShell (`watcher-cp850.ps1`) est modifié, il est nécessaire de redéployer le service :

```powershell
CP850WatcherService.exe uninstall
CP850WatcherService.exe install
CP850WatcherService.exe start
