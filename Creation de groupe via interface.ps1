# Déconnexion préalable
Disconnect-MgGraph -ErrorAction SilentlyContinue

# Chargement des bibliothèques
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms
Import-Module ActiveDirectory -ErrorAction SilentlyContinue

# Connexion à Microsoft Graph
try {
    Connect-MgGraph -Scopes "Directory.Read.All", "Group.ReadWrite.All", "GroupMember.ReadWrite.All", "User.Read.All"
    Write-Host "✅ Connecté en tant que : $((Get-MgContext).Account)"
} catch {
    [System.Windows.Forms.MessageBox]::Show("❌ Échec de la connexion à Microsoft Graph.`n$($_.Exception.Message)", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Vérification d'accès
$user = Get-MgUser -UserId (Get-MgContext).Account
$allowedUsers = @("adelp@2fttr0.onmicrosoft.com", "n-f@2fttr0.onmicrosoft.com")
if ($allowedUsers -notcontains $user.UserPrincipalName) {
    [System.Windows.Forms.MessageBox]::Show("⛔ Vous n'êtes pas autorisé à utiliser ce programme.", "Accès refusé", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}

# Connexion à SharePoint
try {
    Connect-PnPOnline -Url "https://2fttr0.sharepoint.com" `
                      -ClientId "30cede37-891b-4cd8-a202-b940dbabbd8f" `
                      -Tenant "01faf909-de94-4fb2-a307-63c61281cd49" `
                      -Interactive
} catch {
    [System.Windows.Forms.MessageBox]::Show("❌ Échec de la connexion à SharePoint.`n$($_.Exception.Message)", "Erreur", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Error)
    exit
}
# Vérification / création des colonnes nécessaires dans les listes
Write-Host "🔍 Vérification des colonnes SharePoint..."

# Vérifier que la liste AllGroup contient bien la colonne GroupId
$allGroupFields = Get-PnPField -List "AllGroup"
if ($allGroupFields.InternalName -notcontains "GroupId") {
    Write-Host "➕ Ajout de la colonne GroupId dans AllGroup..."
    Add-PnPField -List "AllGroup" -DisplayName "GroupId" -InternalName "GroupId" -Type Text -AddToDefaultView
}

# Vérifier que la liste AllUser contient bien la colonne Email
$allUserFields = Get-PnPField -List "AllUser"
if ($allUserFields.InternalName -notcontains "Email") {
    Write-Host "➕ Ajout de la colonne Email dans AllUser..."
    Add-PnPField -List "AllUser" -DisplayName "Email" -InternalName "Email" -Type Text -AddToDefaultView
}

Write-Host "✅ Colonnes vérifiées."
# -------------------------------
# 🔄 Synchronisation des listes
# -------------------------------

Write-Host "🔄 Synchronisation de la liste AllUser..."
$allAADUsers = Get-MgUser -All | Select-Object DisplayName, UserPrincipalName
$allSPUsers = Get-PnPListItem -List "AllUser"

# Supprimer les utilisateurs disparus
foreach ($spUser in $allSPUsers) {
    $email = $spUser["Email"]
    if ($allAADUsers.UserPrincipalName -notcontains $email) {
        Remove-PnPListItem -List "AllUser" -Identity $spUser.Id -Force
    }
}

# Ajouter les nouveaux utilisateurs
foreach ($aadUser in $allAADUsers) {
    if (-not ($allSPUsers | Where-Object { $_["Email"] -eq $aadUser.UserPrincipalName })) {
        Add-PnPListItem -List "AllUser" -Values @{
            Title = $aadUser.DisplayName
            Email = $aadUser.UserPrincipalName
        }
    }
}

Write-Host "✅ Liste AllUser synchronisée."

Write-Host "🔄 Synchronisation de la liste AllGroup..."
$allAADGroups = Get-MgGroup -All | Select-Object DisplayName, Id
$allSPGroups = Get-PnPListItem -List "AllGroup"

# Supprimer les groupes disparus
foreach ($spGroup in $allSPGroups) {
    $gid = $spGroup["GroupId"]   # <-- Correction ici
    if ($allAADGroups.Id -notcontains $gid) {
        Remove-PnPListItem -List "AllGroup" -Identity $spGroup.Id -Force
    }
}

# Ajouter les nouveaux groupes
foreach ($aadGroup in $allAADGroups) {
    if (-not ($allSPGroups | Where-Object { $_["GroupId"] -eq $aadGroup.Id })) {   # <-- Correction ici
        Add-PnPListItem -List "AllGroup" -Values @{
            Title   = $aadGroup.DisplayName
            GroupId = $aadGroup.Id             # <-- Correction ici
        }
    }
}

Write-Host "✅ Liste AllGroup synchronisée."

# -------------------------------
# 📥 Récupération des données
# -------------------------------
$entreprises = Get-PnPListItem -List "Test_groupe" | ForEach-Object {
    [PSCustomObject]@{ Nom = $_["Title"]; Abrev = $_["Description"] }
}
$pays = Get-PnPListItem -List "ListePaysFINALvrm" | ForEach-Object {
    [PSCustomObject]@{ Nom = $_["Title"]; Abrev = $_["Abreviation"] }
}
$departements = Get-PnPListItem -List "Abreviation_Departement" | ForEach-Object {
    [PSCustomObject]@{ Nom = $_["Title"]; Abrev = $_["Abreviation"] }
}
$utilisateurs = Get-PnPListItem -List "AllUser" | ForEach-Object {
    [PSCustomObject]@{ NomComplet = $_["Title"]; Email = $_["Email"] }
}
$groupes = Get-PnPListItem -List "AllGroup" | ForEach-Object {
    [PSCustomObject]@{ Nom = $_["Title"]; Id = $_["GroupId"] }   # <-- Correction ici
}

# -------------------------------
# 🖼️ Interface graphique
# -------------------------------

$entrepriseMap = @{}
$paysMap = @{}
$departementMap = @{}
$membreMap = @{}
$groupeMap = @{}

$form = New-Object System.Windows.Forms.Form
$form.Text = "THE CREATOR VFINAL"
$form.Size = New-Object System.Drawing.Size(520, 720)
$form.StartPosition = "CenterScreen"
$form.BackColor = [System.Drawing.Color]::FromArgb(0,5,0)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

function Create-Dropdown($items, $labelText, $y, [ref]$map, $groupBox) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $labelText
    $label.Location = New-Object System.Drawing.Point(10, $y)
    $label.Size = New-Object System.Drawing.Size(100, 20)
    $label.ForeColor = [System.Drawing.Color]::White
    $groupBox.Controls.Add($label)

    $dropdown = New-Object System.Windows.Forms.ComboBox
    $dropdown.Location = New-Object System.Drawing.Point(120, $y)
    $dropdown.Size = New-Object System.Drawing.Size(350, 20)
    $dropdown.DropDownStyle = "DropDownList"
    $dropdown.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
    $dropdown.ForeColor = [System.Drawing.Color]::White

    foreach ($item in $items) {
        if ($null -ne $item -and $item.Nom -and $item.Abrev) {
            $dropdown.Items.Add($item.Nom)
            $map.Value[$item.Nom] = $item.Abrev
        }
    }

    $groupBox.Controls.Add($dropdown)
    return $dropdown
}

# GroupBox infos
$gbSelection = New-Object System.Windows.Forms.GroupBox
$gbSelection.Text = "Informations du groupe"
$gbSelection.Location = New-Object System.Drawing.Point(10, 10)
$gbSelection.Size = New-Object System.Drawing.Size(480, 210)
$gbSelection.ForeColor = [System.Drawing.Color]::White
$gbSelection.BackColor = [System.Drawing.Color]::FromArgb(30,30,30)
$form.Controls.Add($gbSelection)

$cbEntreprise = Create-Dropdown $entreprises "Produit :" 20 ([ref]$entrepriseMap) $gbSelection
$cbDepartement = Create-Dropdown $departements "Département :" 60 ([ref]$departementMap) $gbSelection
$cbPays = Create-Dropdown $pays "Pays :" 100 ([ref]$paysMap) $gbSelection

# Membres
$labelMembre = New-Object System.Windows.Forms.Label
$labelMembre.Text = "Membre :"
$labelMembre.Location = New-Object System.Drawing.Point(10, 140)
$labelMembre.Size = New-Object System.Drawing.Size(100, 20)
$labelMembre.ForeColor = [System.Drawing.Color]::White
$gbSelection.Controls.Add($labelMembre)

$cbMembre = New-Object System.Windows.Forms.ComboBox
$cbMembre.Location = New-Object System.Drawing.Point(120, 140)
$cbMembre.Size = New-Object System.Drawing.Size(350, 20)
$cbMembre.DropDownStyle = "DropDownList"
$cbMembre.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
$cbMembre.ForeColor = [System.Drawing.Color]::White

foreach ($user in $utilisateurs) {
    $cbMembre.Items.Add($user.NomComplet)
    $membreMap[$user.NomComplet] = $user.Email
}
$gbSelection.Controls.Add($cbMembre)

# Groupe source
$labelGroupeSource = New-Object System.Windows.Forms.Label
$labelGroupeSource.Text = "Fusion avec groupe :"
$labelGroupeSource.Location = New-Object System.Drawing.Point(10, 170)
$labelGroupeSource.Size = New-Object System.Drawing.Size(120, 20)
$labelGroupeSource.ForeColor = [System.Drawing.Color]::White
$gbSelection.Controls.Add($labelGroupeSource)

$cbGroupeSource = New-Object System.Windows.Forms.ComboBox
$cbGroupeSource.Location = New-Object System.Drawing.Point(140, 170)
$cbGroupeSource.Size = New-Object System.Drawing.Size(330, 20)
$cbGroupeSource.DropDownStyle = "DropDownList"
$cbGroupeSource.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
$cbGroupeSource.ForeColor = [System.Drawing.Color]::White

foreach ($grp in $groupes) {
    $cbGroupeSource.Items.Add($grp.Nom)
    $groupeMap[$grp.Nom] = $grp.Id
}
$gbSelection.Controls.Add($cbGroupeSource)

# Zone description
$descLabel = New-Object System.Windows.Forms.Label
$descLabel.Text = "Description :"
$descLabel.Location = New-Object System.Drawing.Point(20, 230)
$descLabel.Size = New-Object System.Drawing.Size(140, 20)
$descLabel.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($descLabel)

$descBox = New-Object System.Windows.Forms.TextBox
$descBox.Location = New-Object System.Drawing.Point(160, 230)
$descBox.Size = New-Object System.Drawing.Size(310, 20)
$descBox.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
$descBox.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($descBox)

# Zone résultat
$resultLabel = New-Object System.Windows.Forms.Label
$resultLabel.Text = "Prévisualisation :"
$resultLabel.Location = New-Object System.Drawing.Point(20, 260)
$resultLabel.Size = New-Object System.Drawing.Size(450, 120)
$resultLabel.BorderStyle = "Fixed3D"
$resultLabel.BackColor = [System.Drawing.Color]::FromArgb(45,45,45)
$resultLabel.ForeColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($resultLabel)

# Boutons
$btnCreer = New-Object System.Windows.Forms.Button
$btnCreer.Text = "Créer le groupe"
$btnCreer.Location = New-Object System.Drawing.Point(170, 400)
$btnCreer.Size = New-Object System.Drawing.Size(150, 30)
$btnCreer.BackColor = [System.Drawing.Color]::DarkSlateGray
$btnCreer.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($btnCreer)

$btnFermer = New-Object System.Windows.Forms.Button
$btnFermer.Text = "Fermer"
$btnFermer.Location = New-Object System.Drawing.Point(330, 400)
$btnFermer.Size = New-Object System.Drawing.Size(100, 30)
$btnFermer.BackColor = [System.Drawing.Color]::DarkRed
$btnFermer.ForeColor = [System.Drawing.Color]::White
$btnFermer.Add_Click({ $form.Close() })
$form.Controls.Add($btnFermer)

# Image
$pictureBox = New-Object System.Windows.Forms.PictureBox
$pictureBox.Location = New-Object System.Drawing.Point(10, 450)
$pictureBox.Size = New-Object System.Drawing.Size(480, 150)
$pictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
$imgPath = Join-Path $PSScriptRoot "C:\Users\adel.pecoraro-ammar\Downloads\eb5dc566-6e30-4b86-b8a7-18cc7f277129.png"   # <-- Correction ici
if (Test-Path $imgPath) {
    $pictureBox.Image = [System.Drawing.Image]::FromFile($imgPath)
}
$form.Controls.Add($pictureBox)

$btnCreer.Add_Click({
    $nomEntreprise = $cbEntreprise.SelectedItem
    $nomDepartement = $cbDepartement.SelectedItem
    $nomPays = $cbPays.SelectedItem
    $description = $descBox.Text
    $nomMembre = $cbMembre.SelectedItem
    $mailMembre = $membreMap[$nomMembre]
    $nomGroupeSource = $cbGroupeSource.SelectedItem
    $idGroupeSource = $groupeMap[$nomGroupeSource]

    if (-not $nomEntreprise -or -not $nomDepartement -or -not $nomPays) {
        $resultLabel.Text = "❌ Veuillez sélectionner tous les champs."
        return
    }

    $prefix = $entrepriseMap[$nomEntreprise]
    $middle = $departementMap[$nomDepartement]
    $suffix = $paysMap[$nomPays]

    $internalName = "$prefix-$middle-$suffix"
    $displayName = "$nomEntreprise $nomDepartement $nomPays"
    $mailNickname = $internalName.ToLower() -replace '[^a-z0-9]', ''

    $resultLabel.Text = "Prévisualisation : $displayName`r`n(interne : $internalName)"

    $confirm = [System.Windows.Forms.MessageBox]::Show("Confirmer la création du groupe '$internalName' ?", "Confirmation", [System.Windows.Forms.MessageBoxButtons]::YesNo)
    if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }

    try {
        $existingGroup = Get-MgGroup -Filter "displayName eq '$internalName'" -ErrorAction Stop
        if ($existingGroup) {
            $resultLabel.Text += "`r`n❌ Le groupe '$internalName' existe déjà."
            return
        }
    } catch {}

    try {
        $group = New-MgGroup -BodyParameter @{       
            displayName     = $internalName
            description     = $description
            mailEnabled     = $true
            mailNickname    = $mailNickname
            securityEnabled = $true
            groupTypes      = @("Unified")
        }

        # Ajout du membre choisi
        if ($mailMembre) {
            try {
                $user = Get-MgUser -UserId $mailMembre -ErrorAction Stop
                New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $user.Id
                $resultLabel.Text += "`r`n✅ Membre ajouté : $mailMembre"
            } catch {
                $resultLabel.Text += "`r`n⚠️ Erreur pour $mailMembre : $($_.Exception.Message)"
            }
        }

        # Ajout des membres depuis groupe source
        if ($idGroupeSource) {
            try {
                $membresSource = Get-MgGroupMember -GroupId $idGroupeSource
                foreach ($membre in $membresSource) {
                    New-MgGroupMember -GroupId $group.Id -DirectoryObjectId $membre.Id -ErrorAction SilentlyContinue
                    $resultLabel.Text += "`r`n✅ Membre ajouté depuis groupe : $($membre.Id)"
                }
            } catch {
                $resultLabel.Text += "`r`n⚠️ Erreur lors de l'ajout depuis le groupe '$nomGroupeSource' : $($_.Exception.Message)"
            }
        }

        $resultLabel.Text += "`r`n✅ Groupe '$internalName' créé avec succès."
    } catch {
        $resultLabel.Text += "`r`n❌ Erreur : $($_.Exception.Message)"
    }
})

$form.Add_Shown({$form.Activate()})
[System.Windows.Forms.Application]::Run($form)