# ==================================================================================
# Surveille le dossier localPath. Si un fichier correspondant au filtre localFilter 
# est copié ou modifié (d'ou les 2 évènements qui sont créés), celà déclenche un 
# transcodage du fichier UTF8 en CP850 du fichier identifié.
# 
# Pour lancer le script automatiquement, il faut le placer dans un service Windows.
# Voir  https://github.com/winsw/winsw
# ==================================================================================

####################################################################################
# Paramètres
####################################################################################
# Dossier qui sera surveillé :
$watchPath = "C:\Work\IN"
# Dossier dans lequel sera écrit les fichiers transcodés :
$outputPath = "C:\Work\OUT"
# Filtre pour les fichiers à surveiller (création et modification) :
$localFilter = "*.txt"
$logFile = "C:\Scripts\Logs\watcher-cp850.log"
$debugMode = $false     # Mode debug pour afficher les stack traces détaillées
####################################################################################

####################################################################################
# Paramètres de debouncing (anti-rebond)
####################################################################################
$debounceTimeMs = 3000  # 3 secondes de délai pour éviter les traitements multiples
$maxRetries = 3         # Nombre maximum de tentatives si le fichier est verrouillé

####################################################################################
# Variables globales pour le debouncing
####################################################################################
$Global:pendingFiles = @{}
$Global:fileTimers = @{}

####################################################################################
# === FONCTION DE LOGGING ===
####################################################################################
function Global:WriteLog() {
	param($message)
    $timestamp = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
    "$timestamp - $message" | Out-File -Append -FilePath $logFile
    Write-Host "$timestamp - $message"
}

function Global:WriteDebugLog() {
    param($message)
    if ($debugMode) {
        WriteLog -message "🐛 DEBUG: $message"
    }
}

####################################################################################
# === Fonction de vérification si fichier accessible ===
####################################################################################
function Global:IsFileAccessible() {
    param($filePath)
    
    try {
        $fileStream = [System.IO.File]::OpenRead($filePath)
        $fileStream.Close()
        return $true
    }
    catch {
        return $false
    }
}

####################################################################################
# === Fonction transcodage CP850 ===
####################################################################################
function Global:SaveToCP850() {
	param($filePath)
    WriteLog -message "Tentative de transcodage du fichier : $filePath"

    # Vérifier que le fichier existe toujours
    if (-not (Test-Path $filePath)) {
        WriteLog -message "⚠️ Fichier introuvable, abandon du transcodage : $filePath"
        return
    }

	# Test lecture UTF-8
	try {
		$content = Get-Content -Path $filePath -Encoding UTF8 -ErrorAction Stop
	} catch {
		Write-Host "❌ Erreur : Impossible de lire '$filePath' comme UTF-8 (fichier non UTF-8 ou corrompu)."
		exit 1
	}
	
	# Construction du nom du fichier destination (_cp850 avant l’extension)
	$fileName = [System.IO.Path]::GetFileName($filePath) # toto.txt
	$baseName = [System.IO.Path]::GetFileNameWithoutExtension($fileName) # toto
	$extension = [System.IO.Path]::GetExtension($fileName) # .txt
	$outputFileName = "$baseName`_cp850$extension" # toto_cp850.txt
	# Chemin complet du fichier destination
	$destinationPath = Join-Path $outputPath $outputFileName

    # Retry logic si le fichier est verrouillé
    $retryCount = 0
    $success = $false

	while ($retryCount -lt $maxRetries -and -not $success) {
        try {
            # Attendre que le fichier soit accessible
            if (-not (IsFileAccessible -filePath $filePath)) {
                WriteLog -message "⏳ Fichier verrouillé, attente... (tentative $($retryCount + 1)/$maxRetries)"
                Start-Sleep -Seconds 2
                $retryCount++
                continue
            }

            # Lecture UTF-8
            $content = Get-Content -Path $filePath -Encoding UTF8 -ErrorAction Stop

			# Remplacements spécifiques avant encodage
			$content = $content -replace "œ", "oe"
			$content = $content -replace "Œ", "OE"
			$content = $content -replace "€", "E"
			$content = $content -replace "’", "'"
			$content = $content -replace "‘", "'"
			$content = $content -replace '“', '"'
			$content = $content -replace '”', '"'

			# Transcodage en CP850
            [System.IO.File]::WriteAllLines($destinationPath, $content, [System.Text.Encoding]::GetEncoding("ibm850"))
            WriteLog -message "✅ Transcodage CP850 réussi : $destinationPath"
            $success = $true
            
        } catch {
            $retryCount++
            WriteLog -message "❌ Erreur lors du transcodage (tentative $retryCount/$maxRetries) : $($_.Exception.Message)"
			WriteDebugLog -message "Stack trace transcodage : $($_.ScriptStackTrace)"
            if ($retryCount -lt $maxRetries) {
                Start-Sleep -Seconds 2
            } else {
                WriteLog -message "❌ Échec définitif du transcodage après $maxRetries tentatives"
            }
        }
	} # while ($retryCount -lt $maxRetries -and -not $success)
}

####################################################################################
# === Fonction de traitement avec debouncing ===
####################################################################################
function Global:ProcessFileWithDebounce() {
    param($filePath, $eventType)

    $fileName = [System.IO.Path]::GetFileName($filePath)
    
    # Ignorer les fichiers déjà convertis
    if ($fileName -like '*_cp850.*') {
        WriteLog -message "Fichier ignoré (déjà converti) : $filePath"
        return
    }
    
    WriteLog -message "Événement $eventType détecté sur fichier : $filePath"
    
    # Annuler le timer précédent s'il existe
    if ($Global:fileTimers.ContainsKey($filePath)) {
        $Global:fileTimers[$filePath].Dispose()
        WriteDebugLog -message "Timer précédent annulé pour : $filePath"
    }
    
    # Créer un nouveau timer pour ce fichier
    $timer = New-Object System.Timers.Timer
    $timer.Interval = $debounceTimeMs
    $timer.AutoReset = $false
    
	# Action à exécuter quand le timer expire
    # Utiliser Add-Member pour attacher le filePath au timer
    $timer | Add-Member -MemberType NoteProperty -Name "FilePath" -Value $filePath
	
    $action = {
        param($source, $e)
        
        try {
			$currentFilePath = $source.FilePath
            WriteDebugLog -message "⏰ Timer expiré, traitement du fichier : $currentFilePath"
            
			# Vérifier que le filePath n'est pas null
            if ([string]::IsNullOrEmpty($currentFilePath)) {
                WriteLog -message "❌ Erreur : FilePath est null ou vide dans le timer"
                return
            }
			
            # Supprimer le timer de la liste
            if ($Global:fileTimers.ContainsKey($currentFilePath)) {
                $Global:fileTimers.Remove($currentFilePath)
				WriteDebugLog -message "Timer supprimé de la liste pour : $currentFilePath"
            }
            
            # Traitement réel du fichier
            SaveToCP850 -filePath $currentFilePath
            
        } catch {
            WriteLog -message "❌ Erreur dans le timer : $($_.Exception.Message)"
			WriteDebugLog -message "Stack trace : $($_.ScriptStackTrace)"
        } finally {
            # Nettoyer le timer
			WriteDebugLog -message "Nettoyage du timer"
            $source.Dispose()
        }
    }
    
    # Enregistrer l'action du timer
    Register-ObjectEvent -InputObject $timer -EventName Elapsed -Action $action | Out-Null
    
    # Stocker le timer
    $Global:fileTimers[$filePath] = $timer
    
    # Démarrer le timer
    $timer.Start()
    
    WriteDebugLog -message "⏱️ Timer de debouncing démarré pour : $capturedFilePath ($debounceTimeMs ms)"
}

####################################################################################
# === Initialisation du watcher ===
####################################################################################
WriteLog -message "🚀 Initialisation du watcher..."

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.Path = $watchPath
$watcher.IncludeSubdirectories = $false
$watcher.Filter = $localFilter
$watcher.EnableRaisingEvents = $true
$watcher.NotifyFilter = [System.IO.NotifyFilters]'FileName, LastWrite'

####################################################################################
# === Abonnement aux événements ===
####################################################################################
Register-ObjectEvent $watcher Created -Action {
    $fullPath = $Event.SourceEventArgs.FullPath
    ProcessFileWithDebounce -filePath $fullPath -eventType "Created"
} | Out-Null

Register-ObjectEvent $watcher Changed -Action {
    $fullPath = $Event.SourceEventArgs.FullPath
    ProcessFileWithDebounce -filePath $fullPath -eventType "Changed"
} | Out-Null

# Optionnel : gérer les renommages
Register-ObjectEvent $watcher Renamed -Action {
    $fullPath = $Event.SourceEventArgs.FullPath
    ProcessFileWithDebounce -filePath $fullPath -eventType "Renamed"
} | Out-Null

WriteLog -message "📋 Événements enregistrés :"
Get-EventSubscriber | Format-Table SubscriptionId, SourceIdentifier, EventName, State

####################################################################################
# === Boucle principale avec nettoyage ===
####################################################################################
try {
    WriteLog -message "🎯 Surveillance en cours sur '$watchPath' (filtre: $localFilter)"
    WriteLog -message "📤 Fichiers de sortie dans '$outputPath'"
    WriteLog -message "⏱️ Anti-rebond configuré à $debounceTimeMs ms"
    WriteLog -message "🛑 CTRL+C pour quitter..."
    
	while ($true) {
        Start-Sleep -Seconds 10
        
        # Nettoyage périodique des timers expirés (optionnel)
        try {
            $expiredTimers = @()
            
            # Utiliser GetEnumerator() avec gestion d'exception pour éviter les modifications concurrentes
            $enumerator = $Global:fileTimers.GetEnumerator()
            
            while ($enumerator.MoveNext()) {
                $key = $enumerator.Current.Key
                $timer = $enumerator.Current.Value
                
                if ($timer -ne $null -and -not $timer.Enabled) {
                    $expiredTimers += $key
                }
            }
            
            if ($expiredTimers.Count -gt 0) {
                WriteDebugLog -message "Nettoyage de $($expiredTimers.Count) timer(s) expiré(s)"
                
                foreach ($key in $expiredTimers) {
                    if ($Global:fileTimers.ContainsKey($key)) {
                        try {
                            $Global:fileTimers[$key].Dispose()
                        } catch {
                            WriteDebugLog -message "Erreur lors de la suppression du timer : $($_.Exception.Message)"
                        }
                        $Global:fileTimers.Remove($key)
                    }
                }
				
				WriteDebugLog -message "✅ Nettoyage terminé avec succès - $($expiredTimers.Count) timer(s) supprimé(s)"
				
            }
        } catch {
            # Si l'énumération échoue à cause de modifications concurrentes, on ignore silencieusement
            WriteDebugLog -message "Nettoyage des timers reporté (modification concurrente détectée)"
        }
    } # while
} finally {
    WriteLog -message "🧹 Nettoyage en cours..."
    
    # Nettoyer tous les timers
    foreach ($timer in $Global:fileTimers.Values) {
        try {
            $timer.Dispose()
        } catch {}
    }
    $Global:fileTimers.Clear()
    
    # Nettoyer les event subscribers
    Get-EventSubscriber | Unregister-Event -Force
    
    WriteLog -message "🧹 Nettoyage des watchers et timers effectué."
}

# EOF
