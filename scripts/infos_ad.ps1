param($user)



# Chargement des assembly pour la partie graphique, elles se situent dans le dossier assembly
# IMPERATIVEMENT DANS LE MEME DOSSIER QUE LE SCRIPT
[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignThemes.Wpf.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignColors.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll') | out-null  
[System.Reflection.Assembly]::LoadFrom('assembly\Dragablz.dll') | out-null  

function LoadXaml ($filename) {    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

# Charge le fichier XAML grâce à la fonction LoadXAML
$Xaml = LoadXaml(".\views\infos_ad.xaml")

$XAML.MetroWindow.RemoveAttribute('x:Class')         # Remplace les élements C# du XAML
$XAML.MetroWindow.RemoveAttribute('mc:Ignorable')
$Reader = (New-Object System.Xml.XmlNodeReader $Xaml) # Remplace les élements C# du XAML
$form = [Windows.Markup.XamlReader]::Load($Reader)
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) } #Attribue une variable à tout les enfants de $form


function Get-ADUsersLastLogon($user)
{
  $dcs = Get-ADDomainController -Filter {Name -like "*HEB*"}
  $users = Get-ADUser $user
  $time = 0  
  

  foreach($user in $users)
  {
    foreach($dc in $dcs)
    { 
      $hostname = $dc.HostName
      $currentUser = Get-ADUser $user.SamAccountName | Get-ADObject -Server $hostname -Properties lastLogonTimeStamp

      if($currentUser.lastLogonTimeStamp -gt $time) 
      {
        $time = $currentUser.lastLogonTimeStamp
      }
    }

    $dt = [DateTime]::FromFileTime($time)
   
    $row =  get-date $dt -Format dd/MM/yyyy

    
    echo $row

    $time = 0
  }
}

$lastconnexion = Get-ADUsersLastLogon $user


$ad_user = get-aduser $user -Properties *

$txt_user.text = $ad_user.samaccountname
$txt_name.text = $ad_user.name
$txt_ou.text = $ad_user.distinguishedname

$date = (Get-ADUser -filter {samaccountname -eq $ad_user.samaccountname} –Properties "DisplayName", "msDS-UserPasswordExpiryTimeComputed" |
Select-Object -Property "Displayname",@{Name="ExpiryDate";Expression={[datetime]::FromFileTime($_."msDS-UserPasswordExpiryTimeComputed")}}).ExpiryDate | get-date -Format dd/MM/yyyy
$txt_expire.text = $date
$txt_edit.text = $ad_user.WhenChanged | get-date -Format dd/MM/yyyy
$txt_lastconnexion.text = $lastconnexion

$groups = Get-ADPrincipalGroupMembership $user | select name

foreach($group in $groups){
    $grp = $group.name -replace "_","__"
    $list_group.Items.Add($grp)
}


$form.ShowDialog() | Out-Null