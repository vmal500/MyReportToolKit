# Nom du module : MyReportToolKit.psm1
# Description: 
# Ce module contient des fonctions pour importer et traiter les données des appareils et des travailleurs,
# générer des tableaux de bord et trouver des anomalies dans les données. Les fonctions incluent :
#  - Get-Files : Importer les fichiers de rapport et de liste des travailleurs.
#  - Find-Device : Rechercher des appareils en fonction du nom d'utilisateur, de la clé d'entreprise ou du numéro de série.
#  - Get-Dashboard : Générer un tableau de bord des appareils par emplacement et langue.
#  - Find-Anomalies : Détecter les anomalies dans les données importées.
# Le module utilise des variables globales de script pour stocker les données et les résultats importés.
# Les fonctions fournissent des invites et des messages d'erreur pour guider l'utilisateur.
# Des exemples d'utilisation et des descriptions de paramètres sont fournis dans les messages d'aide -h de chaque fonction.
# Ce module est conçu pour être utilisé dans un environnement PowerShell.
# ATTENTION : Assurez-vous que les fichiers CSV ont le format et la structure corrects comme spécifié dans les descriptions des fonctions.
# Dans le cadre de ce module, les fichiers REPORT et WorkerList sont générés par un autre système, ce qui signifie que les données à l'intérieur de ces fichiers CSV, tant qu'elles répondent aux critères de validation déjà établis, sont considérées comme valides.
# Pour exporter les données stockées dans les variables, utilisez la commande Export-CSV. Par exemple :
#  $script:REPORT | Export-Csv -Path "C:\path\to\export\report.csv" -NoTypeInformation
#  $script:WORKERLIST | Export-Csv -Path "C:\path\to\export\workerlist.csv" -NoTypeInformation
#  $script:DASHBOARD | Export-Csv -Path "C:\path\to\export\dashboard.csv" -NoTypeInformation
#  $script:ANOMALIES | Export-Csv -Path "C:\path\to\export\anomalies.csv" -NoTypeInformation

# Ces variables sont accessibles à toutes les fonctions du module
$script:REPORT = $null
$script:WORKERLIST = $null
$script:foundDevices = $null
$script:DASHBOARD = $null
$script:ANOMALIES = $null

function Get-Files {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$r,
        
        [Parameter(Mandatory=$false)]
        [string]$w,

        [Parameter(Mandatory=$false)]
        [Alias("h")]
        [switch]$Help
    )

    if ($Help) {
        Write-Host @"
Get-Files: Importer les fichiers de rapport et de liste de travailleurs

Usage:
    Get-Files [-r <ReportPath>] [-w <WorkerPath>] [-h]

Paramètres:
    -r  Path vers le fichier CSV de rapport. Si non fourni, vous serez invité à le saisir.
    -w  Path vers le fichier CSV de la liste des travailleurs. Si non fourni, vous serez invité à le saisir.
    -h  Afficher ce message d'aide.

Description:
    Cette fonction importe deux fichiers CSV : un fichier de rapport et un fichier de liste de travailleurs.
    Le fichier de rapport doit avoir un format d'en-tête spécifique, et la liste des travailleurs
    doit contenir des numéros d'identification à 4 chiffres ou moins.

    Si les chemins des fichiers ne sont pas fournis en tant que paramètres, vous serez invité à les saisir.
    Vous pouvez taper 'q' à n'importe quelle invite pour quitter la fonction.

Exemples:
    Get-Files
    Get-Files -r "C:\path\to\report.csv" -w "C:\path\to\workers.csv"
    Get-Files -h
"@
        return
    }

    function Import-Report {
        param ([string]$ReportPath)

        # Définir l'en-tête attendu pour le fichier de rapport
        $expectedHeader = "Device;SerialNumber;DeviceType;DeviceMake;DeviceModel;DiskNb;RAM;Location;LastUptime;LastLogin;WorkerName;BusinessKey;Status;OS"
        $expectedFieldCount = 14  # Nombre de champs dans l'en-tête attendu

        do {
            # Demander à l'utilisateur le chemin du fichier WorkerPath s'il n'est pas fourni
            if (-not $ReportPath) {
                $ReportPath = Read-Host "Entrez le chemin vers le fichier CSV de rapport (ou 'q' pour quitter)"
            }
            # Quitter la fonction si l'utilisateur entre 'q'
            if ($ReportPath -in 'q') {
                Write-Host "Quitter la fonction."
                return $false
            }
            # Vérifier si le fichier existe
            if (-not (Test-Path $ReportPath -PathType Leaf)) {
                Write-Host "Fichier non trouvé. Veuillez réessayer."
                $ReportPath = $null
                continue
            }

            $fileContent = Get-Content $ReportPath

            # Vérifier si le fichier est vide
            if (-not $fileContent) {
                Write-Host "Le fichier est vide. Veuillez fournir un fichier CSV valide."
                $ReportPath = $null
                continue
            }

            # Vérifier l'en-tête
            $actualHeader = $fileContent[0].Trim()
            if ($actualHeader -ne $expectedHeader) {
                Write-Host "Le fichier ne contient pas les en-têtes corrects. Format requis :"
                Write-Host $expectedHeader
                $ReportPath = $null
                continue
            }

            # Vérifier la structure de chaque ligne
            $isValid = $true
            foreach ($line in $fileContent) {
                $fields = $line.Split(';')
                if ($fields.Count -ne $expectedFieldCount) {
                    $isValid = $false
                    break
                }
            }

            if (-not $isValid) {
                Write-Host "Le fichier n'est pas un CSV valide. Chaque ligne doit comporter exactement $expectedFieldCount champs."
                $ReportPath = $null
                continue
            }

            # Si nous sommes arrivés jusqu'ici, le fichier est valide
            $script:REPORT = Import-Csv -Path $ReportPath -Delimiter ";"
            return $true

        } while ($true)
    }

    function Import-Workers {
        param ([string]$WorkerPath)

        do {
            # Demander à l'utilisateur le chemin du fichier WorkerPath s'il n'est pas fourni
            if (-not $WorkerPath) {
                $WorkerPath = Read-Host "Entrez le chemin vers le fichier CSV de la liste des travailleurs (ou 'q' pour quitter)"
            }

            # Quitter la fonction si l'utilisateur entre 'q'
            if ($WorkerPath -in 'q') {
                Write-Host "Quitter la fonction."
                return $false
            }

            # Vérifier si le fichier existe
            if (-not (Test-Path $WorkerPath -PathType Leaf)) {
                Write-Host "Fichier non trouvé. Veuillez réessayer."
                $WorkerPath = $null
                continue
            }

            $fileContent = Get-Content $WorkerPath

            # Vérifier si le fichier est vide
            if (-not $fileContent) {
                Write-Host "Le fichier est vide. Veuillez fournir un fichier CSV valide."
                $WorkerPath = $null
                continue
            }

            # Vérifier la structure du CSV
            $isValid = $true
            foreach ($line in $fileContent) {
                $line = $line.Trim()
                if ($line -and ($line.Contains(';') -or $line.Contains('"'))) {
                    $isValid = $false
                    break
                }
            }

            # Si la structure du CSV est invalide, demander à l'utilisateur de fournir un fichier valide
            if (-not $isValid) {
                Write-Host "Le fichier n'est pas un CSV valide. Veuillez fournir un fichier CSV avec des valeurs séparées par des virgules et sans points-virgules ou guillemets."
                $WorkerPath = $null
                continue
            }

            $workerData = Import-Csv -Path $WorkerPath
            # Vérifie si les valeurs ne correspondent pas à un nombre de 1 à 4 chiffres
            $invalidEntries = $workerData | Where-Object { $_.PSObject.Properties.Value -notmatch '^\d{1,4}$' }

            # Vérifier les entrées invalides dans la liste des travailleurs
            if ($invalidEntries) {
                Write-Host "La liste des travailleurs contient des entrées invalides. Chaque entrée doit être un nombre de 4 chiffres ou moins."
                $WorkerPath = $null
                continue
            }

            # Si le fichier est valide, importer la liste des travailleurs
            $script:WORKERLIST = $workerData
            return $true
        } while ($true)
    }
    $reportImported = if ($r) { Import-Report -ReportPath $r } else { Import-Report }
    if ($reportImported -eq $false) { return }

    $workersImported = if ($w) { Import-Workers -WorkerPath $w } else { Import-Workers }
    if ($workersImported -eq $false) { return }

    if ($reportImported -and $workersImported) {
        Write-Host "Importation réussie, vous pouvez commencer à utiliser les autres fonctions."
    }
}

function Find-Device {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$u,

        [Parameter(Mandatory=$false)]
        [string]$k,

        [Parameter(Mandatory=$false)]
        [string]$s,

        [Parameter(Mandatory=$false)]
        [Alias("h")]
        [switch]$Help
    )

    if ($Help) {
        Write-Host @"
Find-Device: Rechercher des appareils dans les données de rapport importées

Usage:
    Find-Device [-u <username>] [-k <businesskey>] [-s <serial-number>] [-h]

Paramètres:
    -u  Rechercher par username (format : cccNNNN, où c est une lettre et N est un chiffre)
    -k  Rechercher par businesskey (numéro �� 4 chiffres)
    -s  Rechercher par numéro de série (8 caractères)
    -h  Afficher ce message d'aide

Description:
    Cette fonction recherche des appareils dans les données de rapport importées en fonction des critères fournis.
    Si aucune donnée n'est importée, elle vous invitera à utiliser la fonction Get-Files.

    Les résultats de la recherche sont affichés sous forme de tableau et stockés dans la variable $foundDevices.

Exemples:
    Find-Device -u abc1234
    Find-Device -k 5678
    Find-Device -s ABCD1234
    Find-Device -h
"@
        return
    }

    # Vérifier si $script:REPORT est vide
    if (-not $script:REPORT) {
        Write-Host "Aucune donnée de rapport trouvée. Veuillez utiliser la fonction Get-Files pour importer des données."
        if ((Read-Host "Voulez-vous importer des fichiers maintenant ? (Y/N)").ToLower() -eq 'y') {
            Get-Files
        } else {
            return
        }
    }
    # Fonction pour afficher les résultats de la recherche
    # Cette fonction prend en paramètre les résultats de la recherche et les affiche sous forme de tableau.
    # Si aucun résultat n'est trouvé, un message est affiché pour en informer l'utilisateur.
    function Show-Results($results) {
        if ($results) {
            $script:foundDevices = $results
            $results | Format-Table -Property Device, SerialNumber, DeviceType, DeviceMake, DeviceModel, DiskNb, RAM, Location, LastUptime, LastLogin, WorkerName, BusinessKey, Status, OS -AutoSize
        } else {
            Write-Host "Aucun appareil correspondant trouvé."
        }
    }

    # Recherche par nom d'utilisateur
    # Si le paramètre $u est fourni, cette section vérifie d'abord si le format du nom d'utilisateur est valide.
    # Le format attendu est cccNNNN, où c est une lettre et N est un chiffre.
    # Si le format est invalide, un message d'erreur est affiché.
    if ($u) {
        if ($u -notmatch '^[a-zA-Z]{3}\d{4}$') {
            Write-Host "Format de nom d'utilisateur invalide. Veuillez entrer un nom d'utilisateur au format cccNNNN."
            return
        }
        $results = $script:REPORT | Where-Object { $_.LastLogin -eq "$u" }
        Show-Results $results
    }

    # Recherche par BusinessKey
    # Si le paramètre $k est fourni, cette section vérifie d'abord si le format de la clé d'entreprise est valide.
    # Le format attendu est un numéro à 4 chiffres.
    # Si le format est invalide, un message d'erreur est affiché.
    elseif ($k) {
        if ($k -notmatch '^\d{4}$') {
            Write-Host "Format de BusinessKey invalide. Veuillez entrer une clé d'entreprise �� 4 chiffres."
            return
        }
        $results = $script:REPORT | Where-Object { $_.BusinessKey -eq "$k" }
        Show-Results $results
    }

    # Recherche par numéro de série
    # Si le paramètre $s est fourni, cette section vérifie d'abord si le format du numéro de série est valide.
    # Le format attendu est une chaîne de 8 caractères.
    # Si le format est invalide, un message d'erreur est affiché.
    elseif ($s) {
        if ($s -notmatch '^.{8}$') {
            Write-Host "Format de numéro de série invalide. Veuillez entrer un numéro de série de 8 caractères."
            return
        }
        $results = $script:REPORT | Where-Object { $_.SerialNumber -eq $s }
        Show-Results $results
    }

    # Si aucun des paramètres de recherche n'est fourni, un message est affiché pour demander à l'utilisateur de spécifier un paramètre de recherche.
    else {
        Write-Host "Veuillez spécifier un paramètre de recherche : -u pour le nom d'utilisateur, -k pour la clé d'entreprise, ou -s pour le numéro de série."
    }
}

function Get-Dashboard {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [string]$O,
        
        [Parameter(Mandatory=$false)]
        [string]$L,

        [Parameter(Mandatory=$false)]
        [Alias("h")]
        [switch]$Help
    )

    if ($Help) {
        Write-Host @"
Get-Dashboard: Générer un tableau de bord des appareils par emplacement et langue

Usage:
    Get-Dashboard [-O <Office>] [-L <Language>] [-h]

Paramètres:
    -O  Filtrer par emplacement de bureau
    -L  Filtrer par langue (Français, Anglais ou Espagnol)
    -h  Affiche ce message d'aide

Description:
    Cette fonction génère un Dashboard montrant le nombre d'appareils par emplacement et langue.
    Si aucune donnée n'est importée, il vous sera demandé d'utiliser la fonction Get-Files.

    Les emplacements incluent : Calgary, Chili, Concord, Edmonton, Labrador City, Mont-Tremblant, Montreal, 
    MSH, MSH - HUB, Quebec, Rouyn-Noranda, Sept-Iles, Sudbury, Terrace, Toronto, Trail, Val d'Or, 
    Vancouver, et Other.

    Les langues sont déterminées par les deux premiers caractères du champ Device :
    LF = French, LE = English, LS = Spanish

    Les résultats sont affichés sous forme de tableau et stockés dans la variable $DASHBOARD.

Exemples:
    Get-Dashboard
    Get-Dashboard -O Montreal
    Get-Dashboard -L French
    Get-Dashboard -O Vancouver -L English
    Get-Dashboard -h
"@
        return
    }

    # Vérifier si $script:REPORT est vide
    if (-not $script:REPORT) {
        Write-Host "Aucune donn��e de rapport trouvée. Veuillez utiliser la fonction Get-Files pour importer des données."
        if ((Read-Host "Voulez-vous importer des fichiers maintenant ? (Y/N)").ToLower() -eq 'y') {
            Get-Files
        } else {
            return
        }
    }

    # Définir les emplacements prédéfinis
    $locations = @(
        "Calgary", "Chili", "Concord", "Edmonton", "Labrador City", "Mont-Tremblant", "Montreal", "MSH", "MSH - HUB",
        "Quebec", "Rouyn-Noranda", "Sept-Iles", "Sudbury", "Terrace", "Toronto", "Trail", "Val d'Or", "Vancouver", "_Other"
    )

    # Définir les langues et leurs abréviations
    $languages = @{
        "LF" = "French"
        "LE" = "English"
        "LS" = "Spanish"
    }

    # Créer une structure de données pour le dashboard
    $dashboard = @{}
    foreach ($loc in $locations) {
        $dashboard[$loc] = @{
            "French" = 0
            "English" = 0
            "Spanish" = 0
        }
    }

    # Filtrer les données par emplacement si un emplacement est spécifié
    if ($O) {
        if ($O -notin $locations) {
            Write-Host "Emplacement spécifié invalide."
            return
        }
        # Si l'emplacement est spécifié, filtrer les données du rapport pour inclure uniquement cet emplacement
        # ou inclure les emplacements qui ne sont pas dans la liste prédéfinie si l'emplacement spécifié est "_Other"
        $filteredReport = $script:REPORT | Where-Object { $_.Location -eq $O -or ($O -eq "_Other" -and $_.Location -notin $locations) }
    } else {
        # Si aucun emplacement n'est spécifié, utiliser toutes les données du rapport
        $filteredReport = $script:REPORT
    }

    # Filtrer les données par langue si une langue est spécifiée
    if ($L) {
        $validLanguages = $languages.Values
        if ($L -notin $validLanguages) {
            Write-Host "Langue spécifiée invalide."
            return
        }
        # Si une langue est spécifiée, filtrer les données du rapport pour inclure uniquement cette langue
        $filteredReport = $filteredReport | Where-Object { $languages[$_.Device.Substring(0,2)] -eq $L }
    }

    # Populer le dashboard avec les données
    foreach ($device in $filteredReport) {
        # Déterminer la langue de l'appareil à partir des deux premiers caractères du champ Device
        $deviceLang = $languages[$device.Device.Substring(0,2)]
        # Déterminer l'emplacement de l'appareil, utiliser "_Other" si l'emplacement n'est pas dans la liste prédéfinie
        $deviceLoc = if ($device.Location -in $locations) { $device.Location } else { "_Other" }
        # Incrémenter le compteur pour la langue et l'emplacement appropriés dans le dashboard
        $dashboard[$deviceLoc][$deviceLang]++
    }

    # Stocker le dashboard dans la variable $DASHBOARD pour un accès ultérieur
    $script:DASHBOARD = $dashboard

    # Préparer les données pour l'affichage
    $displayData = $dashboard.GetEnumerator() | Sort-Object Name | ForEach-Object {
        $loc = $_.Key
        [PSCustomObject]@{
            Office = $loc
            French = $_.Value.French
            English = $_.Value.English
            Spanish = $_.Value.Spanish
            Total = ($_.Value.French + $_.Value.English + $_.Value.Spanish)
        }
    }

    # Si un bureau spécifique est spécifié, afficher uniquement les données pour ce bureau
    if ($O) {
        $displayData = $displayData | Where-Object { $_.Office -eq $O }
    }

    # Afficher les données du dashboard sous forme de tableau
    $displayData | Format-Table -AutoSize
}

function Find-Anomalies {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$false)]
        [switch]$A,

        [Parameter(Mandatory=$false)]
        [int]$r,
        
        [Parameter(Mandatory=$false)]
        [switch]$L,

        [Parameter(Mandatory=$false)]
        [Alias("h")]
        [switch]$Help
    )

    if ($Help) {
        Write-Host @"
Find-Anomalies: Détecter les anomalies dans les données importées

Usage:
    Find-Anomalies [-A] [-r <rule_number>] [-L] [-h]

Paramètres:
    -A  Exécuter toutes les règles de détection d'anomalies
    -r  Exécuter une règle spécifique par son numéro
    -y  Importer automatiquement les fichiers si aucune donnée n'est présente
    -L  Lister toutes les règles de détection d'anomalies disponibles
    -h  Afficher ce message d'aide

Description:
    Cette fonction recherche des anomalies dans les données importées du rapport et de la liste des travailleurs.
    Si aucune donnée n'est importée, elle vous invitera à utiliser la fonction Get-Files.

    Les anomalies détectées sont stockées dans la variable $ANOMALIES.

Exemples:
    Find-Anomalies -A
    Find-Anomalies -r 1
    Find-Anomalies -L
    Find-Anomalies -h
"@
        return
    }

    # Vérifier si $script:REPORT et $script:WORKERLIST sont vides
    if (-not $script:REPORT -or -not $script:WORKERLIST) {
        Write-Host "Les données du rapport ou de la workers list sont manquantes. Veuillez utiliser la fonction Get-Files pour importer les données."
        if ((Read-Host "Voulez-vous importer des fichiers maintenant ? (Y/N)").ToLower() -eq 'y') {
            Get-Files
        } else {
            return
        }
    }

    # Initialiser $ANOMALIES
    $script:ANOMALIES = @()

    # Définir les règles de détection d'anomalies
    $rules = @(
        @{
            Number = 1
            Description = "Un employé qui a quitté l'entreprise a toujours un ordinateur portable."
            Action = {
                # Récupérer les BusinessKeys de la WORKERLIST
                $workerKeys = $script:WORKERLIST | ForEach-Object { $_.PSObject.Properties.Value }
                # Vérifier chaque entrée du REPORT pour voir si la BusinessKey n'est pas dans la WORKERLIST
                $script:REPORT | Where-Object { $_.BusinessKey -and ($_.BusinessKey -notin $workerKeys) } | ForEach-Object {
                    # Ajouter une anomalie à la liste $ANOMALIES si la BusinessKey n'est pas trouvée dans la WORKERLIST
                    $script:ANOMALIES += [PSCustomObject]@{
                        Rule = 1
                        Description = "BusinessKey non présente dans WORKERLIST"
                        Device = $_.Device
                        BusinessKey = $_.BusinessKey
                    }
                }
            }
        }
        # Ajouter plus de règles
    )

    # Fonction pour lister toutes les règles
    function Show-Rules {
        Write-Host "Règles de détection d'anomalies disponibles :"
        # Parcourir chaque règle et afficher son numéro et sa description
        $rules | ForEach-Object {
            Write-Host "Règle $($_.Number): $($_.Description)"
        }
    }

    # Lister les règles si -L est spécifié
    if ($L) {
        Show-Rules
        return
    }
    # Fonction pour exécuter une règle spécifique
    function Invoke-Rule {
        param (
            [string]$ruleNumber
        )
        
        while ($true) {
            try {
                # Valider que $ruleNumber est bien un numéro
                if (-not [int]::TryParse($ruleNumber, [ref]$null)) {
                    throw "Le numéro de règle spécifié n'est pas valide."
                }
                
                # Trouver la règle correspondant au numéro spécifié
                $rule = $rules | Where-Object { $_.Number -eq [int]$ruleNumber }
                if ($rule) {
                    Write-Host "Exécution de la règle $($rule.Number): $($rule.Description)"
                    # Exécuter l'action associée à la règle
                    & $rule.Action
                    break
                } else {
                    Write-Host "Numéro de règle $ruleNumber non trouvé."
                    break
                }
            } catch {
                Write-Host "$_. Veuillez entrer un numéro ou 'q' pour quitter."
                $ruleNumber = Read-Host "Entrez un numéro de règle"
                if ($ruleNumber -eq 'q') {
                    return
                }
            }
        }
    }

    # Exécuter les règles en fonction des paramètres
    if ($A) {
        Write-Host "Exécution de toutes les règles de détection d'anomalies..."
        # Parcourir chaque règle et l'exécuter
        foreach ($rule in $rules) {
            Invoke-Rule $rule.Number
        }
    } elseif ($r) {
        # Exécuter une règle spécifique si le paramètre -r est spécifié
        Invoke-Rule $r
    } else {
        Write-Host "Veuillez spécifier -A pour exécuter toutes les règles, -r <numéro> pour exécuter une règle spécifique, ou -L pour lister toutes les règles."
        return
    }

    # Afficher les résultats
    if ($script:ANOMALIES.Count -gt 0) {
        Write-Host "Anomalies détectées :"
        # Afficher les anomalies détectées sous forme de tableau
        $script:ANOMALIES | Format-Table -AutoSize
    } else {
        Write-Host "Aucune anomalie détectée."
    }
}