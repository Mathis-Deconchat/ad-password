[CmdletBinding()]
param (    
    [Parameter()]
    [string]
    $message
)

[System.Reflection.Assembly]::LoadWithPartialName('presentationframework') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\System.Windows.Interactivity.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MahApps.Metro.dll')      | out-null  
[System.Reflection.Assembly]::LoadFrom('assembly\LiveCharts.dll')        | out-null  	
[System.Reflection.Assembly]::LoadFrom('assembly\LiveCharts.Wpf.dll') 	 | out-null  
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignThemes.Wpf.dll') | out-null
[System.Reflection.Assembly]::LoadFrom('assembly\MaterialDesignColors.dll') | out-null


#########################################################################
#                        Fonction pour charger les fichiers xaml        #
#########################################################################
# Prend un fichier en entrée, et le charge dans un objet de type System.Xml.XmlDocument pui retourne cet objet

function LoadXaml ($filename) {    
    $XamlLoader = (New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}

#########################################################################
#          Vérification des fichiers/ Dossiers nécessaire               #
#########################################################################



#########################################################################
#          Reader XAML                                                  #
#########################################################################

$Xaml = LoadXaml(".\views\message.xaml")
$XAML.MetroWindow.RemoveAttribute('x:Class')         # Remplace les élements C# du XAML
$XAML.MetroWindow.RemoveAttribute('mc:Ignorable')
$Reader = (New-Object System.Xml.XmlNodeReader $Xaml) # Remplace les élements C# du XAML
$form = [Windows.Markup.XamlReader]::Load($Reader)
$xaml.SelectNodes("//*[@Name]") | ForEach-Object { Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name) } #Attribue une variable à tout les enfants de $form


$icon.kind = "CheckBold"
$msg_content.text = $message


$but_message_ok.Add_Click( { 
        $form.Close() 
        exit ;
    })


$form.showDialog() | out-null
