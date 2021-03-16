[Environment]::CurrentDirectory = "M:\ShadowScriptv1\admin"

# Chargement des assembly pour la partie graphique, elles se situent dans le dossier assembly

[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignThemes.Wpf.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignColors.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll') | out-null  
[System.Reflection.Assembly]::LoadFrom('assembly\Dragablz.dll') | out-null  

# Crï¿½er le dossier de log ainsi que le fichier de log
if (!(test-path $env:userprofile\log)) {
    new-item -path $env:userprofile\log -ItemType Directory | Out-Null
    new-item -path $env:userprofile\log\log_pem.txt -ItemType File | Out-Null
}

# Log la connexion de utilisateur
$dat = get-date -format "yyyy/MM/dd HH:mm"
$log = "[$dat]//: $env:username start app..."
Add-Content -Value $log -Path $env:userprofile\log\log_pem.txt -Force

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
# REINIT_MDP_AD
##########
function infos_ad {
    $user = $($List_mdp_ad.SelectedItem -split " -")[0] 
    $user = $user -replace "__", "_"
    Start-Process powershell.exe -argument " .\scripts\infos_ad.ps1 -user $user" -WindowStyle Hidden
}

##########
# REINIT_MDP_AD
##########
## Rï¿½initialise le mot de passe AD
function REINIT_MDP_AD {
    $dom = $null
    # Sï¿½lectionne le nom d'utilisateur dans la liste (prends uniquement le nom user car liste est composï¿½ de {username} - {nom complet}
    $user = $($List_mdp_ad.SelectedItem -split " -")[0] 
    $user = $user -replace "__", "_"
    Add-Type -AssemblyName 'System.Web' 
    ## Création du mot de passe aléatoire - API motdepasse.xyz - 12 charactères : chiffres / MAJ / MINUSCULES / caractères similaires exclus
    $Pass = Invoke-RestMethod -Uri "https://api.motdepasse.xyz/create/?include_digits&include_lowercase&include_uppercase&include_special_characters&exclude_similar_characters&password_length=12&quantity=1"
    $Password = [string]$Pass.passwords
    $Password = $Password -replace '<', '@' -replace '>', '_' 
    # Set le password
    Set-ADAccountPassword -Identity $user -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $Password -Force )
    write-host $user
    
    $CurrentDomain = 'LDAP://' + ([ADSI]"").distinguishedName
    $dom = New-Object System.DirectoryServices.DirectoryEntry($CurrentDomain, $User, $Password) 
    $res = $dom.path


    if ( $res -ne $null) {  
        
        $Message = 'Bonjour,


Votre mot de passe a été réinitialisé à : '+ $password + '


Afin de le modifier, après vous être connecté au serveur, utiliser le raccourci "Chgt MDP Agile" présent sur votre bureau.


Puis sélectionner l''option "Modifier un mot de passe".


Il doit répondre aux conditions suivantes :

- Une longueur minimum de 12 caractères avec au moins une minuscule, une majuscule, un chiffre et un caractère spécial.

- Impossibilité de réutiliser les 5 derniers mots de passe.


Nous restons à votre disposition pour tout renseignement complémentaire.


Bien cordialement,
 '
    
        $OutputEncoding = (New-Object System.Text.UnicodeEncoding $False, $False).psobject.BaseObject
        $Message | clip
        Start-Process powershell.exe -argument " .\scripts\message.ps1 -ok ok -user $user" -NoNewWindow

        $dat = get-date -format "yyyy/MM/dd HH:mm:ss"
        $log = "[$dat]//: [TECH: $env:username] a réintialiser le mot de passe de [USER: $user]."
        Add-Content -Value $log -Path $env:userprofile\log\log_pem.txt -Force  
    }


}

#_______________________#
## TAB AD
#_______________________#
# Click sur le bouton de rï¿½intialisation de l'AD
$user = $null ;
$GLOBAL:user_ad = get-aduser -filter { Enabled -eq $true }  

$chk_inactif_agile.IsChecked = ''

$chk_inactif_agile.Add_UnChecked( {
        $GLOBAL:user_ad = Get-ADUser -Filter { enabled -eq $true } -Properties * | ? { ($_.AccountExpirationDate -eq $NULL -or $_.AccountExpirationDate -gt (Get-Date)) } 
    })
$chk_inactif_agile.Add_Checked( {
        $GLOBAL:user_ad = Get-ADUser -Filter { enabled -eq $false } -Properties * | ? { ($_.AccountExpirationDate -eq $NULL -or $_.AccountExpirationDate -lt (Get-Date)) } 
    })
 
$Recherche_mdp_ad.Add_TextChanged( {     
        $list_mdp_ad.Items.Clear()
        $GLOBAL:unique_user = $user_ad | ? { $_.Name -like "*$($Recherche_mdp_ad.text)*" -or $_.samaccountname -like "*$($Recherche_mdp_ad.text)*" }

        foreach ($a in $unique_user) {  
            $cleanUser = $a.samaccountname -replace "_", "__"
            $item_ad = $cleanUser + " - " + $a.name
                          
            $list_mdp_ad.Items.Add("$item_ad")
            #$list_mdp_ad.Items.Add.Value('pomme')
        
        }
    
    }) # FIN RECHERCHE AD TEXT CHANGED

$but_reinit_ad.Add_Click( {
        REINIT_MDP_AD
    })

$but_infos_ad.Add_Click( {
        INFOS_AD
    })




$form.showDialog() | out-null