param (
    [string]$JsonPath   = "C:\Temp\groupData.json",
    [string]$StatusPath = "C:\Temp\creationStatus.txt"
)

# =========================
# CONFIGURATION GRAPH
# =========================
$ApplicationId   = "??"
$ClientSecret    = "??"
$TenantId        = "??"

$SecureSecret = ConvertTo-SecureString $ClientSecret -AsPlainText -Force
$Credential   = New-Object System.Management.Automation.PSCredential ($ApplicationId, $SecureSecret)

# =========================
# INIT
# =========================
Disconnect-Graph -ErrorAction SilentlyContinue
if (-not (Test-Path "C:\Temp")) { New-Item -ItemType Directory -Path "C:\Temp" -Force | Out-Null }

$CRE = 0
$ErrorMessage = ""
Set-Content -Path $StatusPath -Value "CRE=0`nERROR="

# =========================
# CONNEXION GRAPH
# =========================
try {
    Connect-MgGraph -TenantId $TenantId -ClientSecretCredential $Credential
    $CurrentUser = (Get-MgContext).Account
}
catch {
    Set-Content $StatusPath "CRE=0`nERROR=Connexion Graph échouée : $($_.Exception.Message)"
    return
}

# =========================
# LECTURE JSON
# =========================
if (-not (Test-Path $JsonPath)) {
    Set-Content $StatusPath "CRE=0`nERROR=Fichier JSON introuvable"
    return
}

try {
    $GroupData = Get-Content $JsonPath -Raw | ConvertFrom-Json
}
catch {
    Set-Content $StatusPath "CRE=0`nERROR=JSON invalide : $($_.Exception.Message)"
    return
}

# =========================
# CRÉATION GROUPE
# =========================
try {
    $GroupParams = @{
        displayName     = $GroupData.displayName
        description     = $GroupData.description
        mailEnabled     = $true
        securityEnabled = $false
        mailNickname    = $GroupData.mailNickname
        groupTypes      = @("Unified")
    }

    $Group = New-MgGroup -BodyParameter $GroupParams
    $CRE = 1
}
catch {
    Set-Content $StatusPath "CRE=0`nERROR=Création groupe échouée : $($_.Exception.Message)"
    return
}

# =========================
# CONVERSION DYNAMIQUE
# =========================
try {
    function ConvertStaticGroupToDynamic {
        param (
            [string]$GroupId,
            [string]$Rule
        )

        $DynamicType = "DynamicMembership"
        [System.Collections.ArrayList]$Types = (Get-MgGroup -GroupId $GroupId).GroupTypes

        if ($Types -contains $DynamicType) {
            throw "Le groupe est déjà dynamique"
        }

        $Types.Add($DynamicType)

        Update-MgGroup `
            -GroupId $GroupId `
            -GroupTypes $Types.ToArray() `
            -MembershipRuleProcessingState "On" `
            -MembershipRule $Rule
    }

    ConvertStaticGroupToDynamic -GroupId $Group.Id -Rule $GroupData.membershipRule
}
catch {
    $CRE = 0
    $ErrorMessage = "Conversion dynamique échouée : $($_.Exception.Message)"
}

# =========================
# SORTIE
# =========================
Set-Content $StatusPath "CRE=$CRE`nERROR=$ErrorMessage"
Set-Content "C:\Temp\currentUser.txt" $CurrentUser
$btnCreer.Add_Click({

    # =========================
    # RÉCUPÉRATION VALEURS
    # =========================
    $nomEntreprise  = $textEntreprise.Text.Trim()
    $nomSousproduit = $textSousProduit.Text.Trim()
    $nomDepartement = $textDepartement.Text.Trim()
    $nomRoles       = $textRoles.Text.Trim()
    $nomPays        = $textPays.Text.Trim()
    $nomVille       = $textVille.Text.Trim()

    # Abréviations
    $prefix   = $entrepriseMap[$nomEntreprise]
    $sousfix  = $SousProduitMap[$nomSousproduit]
    $depfix   = $departementMap[$nomDepartement]
    $middle   = $rolesMap[$nomRoles]
    $suffix   = $paysMap[$nomPays]
    $villefix = $villeMap[$nomVille]

    # =========================
    # GÉNÉRATION DES NOMS
    # =========================
    function Generate-GroupNames {
        param (
            $ne, $nsp, $np, $nv, $nd, $nr,
            $pfx, $spx, $sfx, $vfx, $dfx, $mfx
        )

        $mail = @($pfx,$spx,$sfx,$vfx,$dfx,$mfx) | Where-Object { $_ }
        $mail += "ALL_EMPLOYEES"
        $mailNick = ($mail -join "_").ToUpper()

        $disp = @($ne,$nsp,$np,$nv,$nd,$nr) | Where-Object { $_ }
        $disp += "ALL-EMPLOYEES"
        $dispName = ($disp -join "-")

        if (($ne -or $nsp) -and ($nd -or $nr) -and ($np -or $nv)) {
            $mailNick = $mailNick -replace "_?ALL_EMPLOYEES",""
            $dispName = $dispName -replace "-?ALL-EMPLOYEES",""
        }

        return @{
            MailNickname = $mailNick
            DisplayName  = $dispName
        }
    }

    $names = Generate-GroupNames `
        $nomEntreprise $nomSousproduit $nomPays $nomVille `
        $nomDepartement $nomRoles `
        $prefix $sousfix $suffix $villefix $depfix $middle

    $mailNickname = ($names.MailNickname -replace "[^A-Z0-9_]", "" -replace "__+","_").Trim("_")
    $displayName  = ($names.DisplayName  -replace "[^A-Za-z0-9\- ]","" -replace "--+","-").Trim("-")

    # =========================
    # CONFIRMATION
    # =========================
    if ([System.Windows.Forms.MessageBox]::Show(
        "Confirmer la création du groupe :`n$displayName",
        "Confirmation",
        "YesNo",
        "Question"
    ) -ne "Yes") { return }

    # =========================
    # RÈGLE DYNAMIQUE
    # =========================
    $rules = @()

    if (-not $extcheckbox.Checked) {
        $rules += '(user.extensionAttribute10 -contains ",treatAsEmployee")'
    }

    if ($manacheckbox.Checked) {
        $rules += '(user.extensionAttribute10 -contains ",ManagerDDL")'
    }

    if ($nomVille) {
        $rules += "(user.extensionAttribute1 -eq `"$nomVille`")"
    }
    elseif ($nomPays) {
        $rules += "(user.co -eq `"$nomPays`")"
    }

    if ($nomRoles) {
        $rules += "(user.extensionAttribute4 -eq `"$nomRoles`")"
    }
    elseif ($nomDepartement) {
        $rules += "(user.extensionAttribute8 -eq `"$nomDepartement`")"
    }

    if ($nomSousproduit) {
        $rules += "(user.extensionAttribute7 -contains `"$nomSousproduit`")"
    }
    elseif ($nomEntreprise) {
        $rules += "(user.extensionAttribute2 -contains `"$nomEntreprise`")"
    }

    $membershipRule = $rules -join " -and "

    # =========================
    # JSON POUR CREATE.PS1
    # =========================
    $json = @{
        displayName    = $displayName
        description    = "Created via Forterro Group Builder"
        mailNickname   = $mailNickname
        membershipRule = $membershipRule
    }

    $json | ConvertTo-Json -Depth 3 | Set-Content "C:\Temp\groupData.json" -Encoding UTF8

    # =========================
    # LANCEMENT CREATE.PS1
    # =========================
    try {
        Start-Process powershell.exe `
            -ArgumentList "-NoProfile -ExecutionPolicy Bypass -File `"C:\Users\adel.pecoraro-ammar\Create.ps1`"" `
            -Wait -WindowStyle Hidden

        $status = Get-Content "C:\Temp\creationStatus.txt" -Raw
        $CRE    = (($status -split "`n")[0] -split "=")[1]
        $ERROR  = (($status -split "`n")[1] -split "=")[1]
        $user   = Get-Content "C:\Temp\currentUser.txt"

        if ($CRE -eq 1) {
            $resultLabel.ForeColor = [System.Drawing.Color]::LightGreen
            $resultLabel.Text = "✅ Groupe créé avec succès`n$displayName`n📧 $mailNickname@forterro.com`n👤 $user"

            Add-PnPListItem -List "log" -Values @{
                Title   = $displayName
                creator = $user
                date    = Get-Date
                rule    = $membershipRule
                mail    = "$mailNickname@forterro.com"
            }
        }
        else {
            $resultLabel.ForeColor = [System.Drawing.Color]::Red
            $resultLabel.Text = "❌ Échec de création`n$displayName`n⚠️ $ERROR"
        }
    }
    catch {
        $resultLabel.ForeColor = [System.Drawing.Color]::Orange
        $resultLabel.Text = "⚠️ Erreur inattendue : $($_.Exception.Message)"
    }
})

