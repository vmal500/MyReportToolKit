# Guide d'usage du module Powershell MyReportToolKit.psm1

## Cas d'utilisation
Ce module est conçu pour les professionnels IT débutants qui doivent gérer et analyser les données relatives aux appareils informatiques d'une entreprise. Il offre des fonctionnalités pour importer des données, rechercher des appareils spécifiques, générer des tableaux de bord et détecter des anomalies dans les données.

## Installation
1. Téléchargez le fichier MyReportToolKit.psm1
2. Placez-le dans un dossier, par exemple :
   `C:\path\to\your\folder\Modules\MyReportToolKit.psm1`
3. Ouvrez PowerShell et exécutez la commande suivante :
   `Import-Module MyReportToolKit.psm1`

## Dépendances
- Windows PowerShell

## Prérequis
- Un fichier CSV de rapport (provenant d'un outil de RMM) dans ce format :
  ```
  Device;SerialNumber;DeviceType;DeviceMake;DeviceModel;DiskNb;RAM;Location;LastUptime;LastLogin;WorkerName;BusinessKey;Status;OS
  LF-KW-XJ0H9EVV;XJ0H9EVV;WINDOWS_LAPTOP;LENOVO;20X5007EUS ThinkPad L14 Gen 2a;1;14.83;Montreal;2024-09-06T15:44:12.000+0000;vim0042;NAME OF USER;;In use;Windows 10 Enterprise Edition
  ...
  ```
- Un fichier CSV des travailleurs avec la liste des numéros d'employés.
  ```
  BusinessKey
  1234
  5678
  ...
  ```

## Table des matières des fonctions

1. [Get-Files](#1-get-files)
2. [Find-Device](#2-find-device)
3. [Get-Dashboard](#3-get-dashboard)
4. [Find-Anomalies](#4-find-anomalies)

### 1. Get-Files

**a. Utilisation :**
Cette fonction importe les fichiers de rapport et de liste des travailleurs.

**c. Description des paramètres :**
- `-r` : Spécifie le chemin du fichier CSV contenant les données du rapport
- `-w` : Spécifie le chemin du fichier CSV contenant la liste des travailleurs
- `-h` : Affiche les instructions d'utilisation de la fonction

**d. Format attendu :**
- `-r` et `-w` : Chaîne de caractères représentant un chemin de fichier valide
- `-h` : Switch (pas de valeur associée)

**e. Résultat attendu :**
Les données sont importées dans les variables globales `$script:REPORT` et `$script:WORKERLIST`
```powershell
Get-Files
Entrez le chemin vers le fichier CSV de rapport (ou 'q' pour quitter): REPORT.csv
Entrez le chemin vers le fichier CSV de la liste des travailleurs (ou 'q' pour quitter): WorkerList.csv                                                       
Importation réussie, vous pouvez commencer à utiliser les autres fonctions.
```
**f. Exemples :**
```powershell
Get-Files -r "C:\rapport.csv" -w "C:\travailleurs.csv"
Get-Files -h
```

### 2. Find-Device

**a. Utilisation :**
Cette fonction recherche des appareils dans les données importées.

**c. Description des paramètres :**
- `-u` : Recherche par nom d'utilisateur
- `-k` : Recherche par clé d'entreprise
- `-s` : Recherche par numéro de série
- `-h` : Affiche les instructions d'utilisation de la fonction

**d. Format attendu :**
- `-u` : Format cccNNNN (3 lettres suivies de 4 chiffres)
- `-k` : Numéro à 4 chiffres ou moins
- `-s` : Chaîne de 8 caractères
- `-h` : Switch (pas de valeur associée)

**e. Résultat attendu :**
Affiche un tableau des appareils correspondant aux critères de recherche
**f. Exemples :**
```powershell
Find-Device -u abc1234
Find-Device -k 5678
```

### 3. Get-Dashboard

**a. Utilisation :**
Cette fonction génère un tableau de bord des appareils par emplacement et langue.

**c. Description des paramètres :**
- `-O` : Filtre les résultats par emplacement de bureau
- `-L` : Filtre les résultats par langue
- `-h` : Affiche les instructions d'utilisation de la fonction

**d. Format attendu :**
- `-O` : Chaîne de caractères correspondant à un emplacement prédéfini. Les emplacements prédéfinis sont stockés dans la variable `$locations` dans le script PowerShell.
- `-L` : "French", "English" ou "Spanish"
- `-h` : Switch (pas de valeur associée)

**e. Résultat attendu :**
Affiche un tableau récapitulatif du nombre d'appareils par emplacement et langue
```powershell
Get-Dashboard                                                                     

Office         French English Spanish Total
------         ------ ------- ------- -----
_Other              4       0       0     4
Calgary             1      46       0    47
Chili               0       4      33    37
Concord             0      10       0    10
Edmonton            0      23       0    23
Labrador City       2       2       0     4
Mont-Tremblant      7       0       0     7
Montreal          280      33       0   313
MSH               337      21       0   358
MSH - HUB          74       4       0    78
Quebec             74       3       0    77
Rouyn-Noranda      22       0       0    22
Sept-Iles           6       0       0     6
Sudbury             0      27       0    27
Terrace             0       4       0     4
Toronto             2      71       0    73
Trail               0      24       0    24
Val d'Or           18       0       0    18
Vancouver           2     123       0   125
```
```powershell
Get-Dashboard -O Vancouver -L English

Office    French English Spanish Total
------    ------ ------- ------- -----
Vancouver      0     123       0   123
```

**f. Exemples :**
```powershell
Get-Dashboard -O Montreal -L French
Get-Dashboard -h
```

### 4. Find-Anomalies

**a. Utilisation :**
Cette fonction détecte les anomalies dans les données importées. Elle est "scalable" car on peut rajouter le nombre de règles que l'on veut avec un peu d'expérience en PowerShell.

**c. Description des paramètres :**
- `-A` : Exécute toutes les règles de détection d'anomalies
- `-r` : Exécute une règle spécifique par son numéro
- `-L` : Affiche la liste de toutes les règles disponibles
- `-h` : Affiche les instructions d'utilisation de la fonction

**d. Format attendu :**
- `-A` : Switch (pas de valeur associée)
- `-r` : Numéro entier correspondant à une règle
- `-L` : Switch (pas de valeur associée)
- `-h` : Switch (pas de valeur associée)

**e. Résultat attendu :**
Affiche les anomalies détectées ou la liste des règles disponibles
```powershell
Find-Anomalies -r 1                             
Exécution de la règle 1: Un employé qui a quitté l'entreprise a toujours un ordinateur portable.
Anomalies détectées :

Rule Description                              Device         BusinessKey
---- -----------                              ------         -----------
   1 BusinessKey non présente dans WORKERLIST LF-KW-PF2SW3SN 5760
```
**f. Exemples :**
```powershell
Find-Anomalies -A
Find-Anomalies -r 1
```

### Validation des entrées

Chaque fonction valide si un rapport ou la liste de travailleur a été chargé pour éviter les erreurs. Aussi, il y a de la validation d'entrée pour éviter les erreurs.
```powershell
Find-Anomalies -r 1                                                               
Les données du rapport ou de la workers list sont manquantes. Veuillez utiliser la fonction Get-Files pour importer les données.
Voulez-vous importer des fichiers maintenant ? (Y/N): 
```
```powershell
Find-Device -k mille trois
Format de nom d'utilisateur invalide. Veuillez entrer un nom d'utilisateur au format cccNNNN.
``` 
