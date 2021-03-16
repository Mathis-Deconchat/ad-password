[Environment]::CurrentDirectory = $PSScriptRoot

# Chargement des assembly pour la partie graphique, elles se situent dans le dossier assembly

[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignThemes.Wpf.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignColors.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll') | out-null  
[System.Reflection.Assembly]::LoadFrom('assembly\Dragablz.dll') | out-null  

# Crï¿½er le dossier de log ainsi que le fichier de log
if (!(test-path $env:userprofile\log)) {
    new-item -path $env:userprofile\log -ItemType Directory | Out-Null
    new-item -path $env:userprofile\log\log_reinit.txt -ItemType File | Out-Null
}

function Log {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $message
    )
    $date = get-date -format "yyyy/MM/dd HH:mm"
    $user = $env:username
    $log = "[$date] : $user // $message"
    Add-Content -Value $log -Path $env:userprofile\log\log_reinit.txt -Force

}

Log -message "Lance l'application"


# Fonction pour charger des fichiers XAML
function LoadXaml ($filename) {    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

# Charge le fichier XAML grï¿½ce ï¿½ la fonction LoadXAML
$Xaml = LoadXaml(".\main.xaml")

$XAML.Window.RemoveAttribute('x:Class')         # Remplace les ï¿½lements C# du XAML
$XAML.Window.RemoveAttribute('mc:Ignorable')
$Reader = (New-Object System.Xml.XmlNodeReader $Xaml) # Remplace les ï¿½lements C# du XAML
$form = [Windows.Markup.XamlReader]::Load($Reader)
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) } #Attribue une variable ï¿½ tout les enfants de $form




##########
# 
##########
function Get-AdInfos {
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $user
    )    
    Start-Process powershell.exe -argument " .\scripts\infos_ad.ps1 -user $user" -WindowStyle Hidden
}

##########
# 
##########
function Set-AdPassword {
    [CmdletBinding()]
    param (
        [Parameter()]
        [String]
        $user
    )
    $dom = $null
    Add-Type -AssemblyName 'System.Web'
    $Pass = Invoke-RestMethod -Uri "https://api.motdepasse.xyz/create/?include_digits&include_lowercase&include_uppercase&include_special_characters&exclude_similar_characters&password_length=12&quantity=1"
    $Password = [string]$Pass.passwords    
    
    Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force )
    
    
    $CurrentDomain = 'LDAP://' + ([ADSI]"").distinguishedName
    $dom = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $User, $Password) 
    $res = $dom.path


    if ( $res -ne $null) { 
        $password  | clip
        $message = "Mot de passe réinitialisé pour $user avec $password `n✔ Copié dans le presse papier "
        Start-Process powershell.exe -argument " .\scripts\message.ps1 -message '$message'" -NoNewWindow        
        Log -message "réintialiser le mot de passe de [USER: $user]."
        
    }
    else {
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Réinitilisation KO", 4, "Erreur de réinitialisation", 0x0 + 16)
    }


}

$user = $null ;
$user_ad = get-aduser -filter { Enabled -eq $true }  

$Recherche_mdp_ad.Add_TextChanged( {     
        $list_mdp_ad.Items.Clear()
        $unique_user = $user_ad | Where-Object { $_.Name -like "*$($Recherche_mdp_ad.text)*" -or $_.samaccountname -like "*$($Recherche_mdp_ad.text)*" }
        foreach ($a in $unique_user) {  
            $cleanUser = $a.samaccountname -replace "_", "__"
            $item_ad = $cleanUser + " - " + $a.name                          
            $list_mdp_ad.Items.Add($item_ad)         
        
        }
    
    }) 

$but_reinit_ad.Add_Click( {
        if ($List_mdp_ad.SelectedItem.length -gt 0) {
            $user = $($List_mdp_ad.SelectedItem -split " -")[0] 
            $user = $user -replace "__", "_"
            Set-AdPassword -user $user
        }
        else {
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("Veuillez sélectionner un utilisateur", 4, "Erreur de sélection", 0x0 + 16)
        }
        
    })

$but_infos_ad.Add_Click( {
        
        if ($List_mdp_ad.SelectedItem.length -gt 0) {
            $user = $($List_mdp_ad.SelectedItem -split " -")[0] 
            $user = $user -replace "__", "_"
            Get-AdInfos -user $user
        }
        else {
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("Veuillez sélectionner un utilisateur", 4, "Erreur de sélection", 0x0 + 16)
        }
        
    })




$form.showDialog() | out-null