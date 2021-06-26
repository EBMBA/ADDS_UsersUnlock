#========================================================================
#
# Tool Name	: ADDS-UnlockUsers
# Author 	: METRAL Emile 
# Gitlab / Github : EBMBA
#
#========================================================================

[Void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
foreach ($item in $(Get-ChildItem .\ -Filter *.dll).name) {
    [Void][System.Reflection.Assembly]::LoadFrom("$item")
}

#########################################################################
#                     Variable + Import Module                          #
#########################################################################
$path = [System.IO.Path]::GetDirectoryName($MyInvocation.MyCommand.Definition)

[void] [Reflection.Assembly]::LoadWithPartialName( 'System.Windows.Forms' )

#Requires -Modules ActiveDirectory
Import-Module ActiveDirectory

#########################################################################
#                        Load Main Panel                                #
#########################################################################

$Global:pathPanel= split-path -parent $MyInvocation.MyCommand.Definition

function LoadXaml ($filename){
    $XamlLoader=(New-Object System.Xml.XmlDocument)
    $XamlLoader.Load($filename)
    return $XamlLoader
}
$XamlMainWindow=LoadXaml($pathPanel+"\main.xaml")
$reader = (New-Object System.Xml.XmlNodeReader $XamlMainWindow)
$Form = [Windows.Markup.XamlReader]::Load($reader)


#########################################################################
#                       Functions Base                   								#
#########################################################################
#region window
$XamlMainWindow.SelectNodes("//*[@Name]") | %{
    try {Set-Variable -Name "$("WPF_"+$_.Name)" -Value $Form.FindName($_.Name) -ErrorAction Stop}
    catch{throw}
    }
 
Function Get-FormVariables{
  if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}
  write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
  get-variable *WPF*
}

#endregion
#########################################################################
#                       Functions                       								#
#########################################################################
function Validate-IsEmptyTrim ([string] $field) {

  if($field -eq $null -or $field.Trim().Length -eq 0) {
    return $true    
  }
      
  return $false
}
#########################################################################
#                       DATA       						       		                #
#########################################################################

$UsersLocks = $(Search-ADAccount -LockedOut | Select-Object -Property Name, SamAccountName)
Write-Host $UsersLocks

$Form.Add_ContentRendered({
  $WPF_Validate.IsEnabled = $false
})

foreach ($UserLock in $UsersLocks) {
  $GroupsList = New-Object PSObject
  $GroupsList = $GroupsList | Add-Member NoteProperty Login $UserLock.SamAccountName -passthru
  $GroupsList = $GroupsList | Add-Member NoteProperty Name $UserLock.Name -passthru	
  $WPF_Users.Items.Add($GroupsList) > $null
}

$WPF_Username.Add_TextChanged({
  try {
    Get-ADUser -Identity $($WPF_Username.Text)
    $WPF_Username.Background = [System.Windows.Media.Brushes]::PaleGreen
    $WPF_Validate.IsEnabled = $True
  }
  catch {
    $WPF_Username.Background = [System.Windows.Media.Brushes]::PaleVioletRed
    $WPF_Validate = $false
  }
})

$WPF_Users.add_SelectionChanged({
  $WPF_Username.Text=$($WPF_Users.SelectedItem).Login
})


$WPF_Validate.Add_Click({
  if (-not(Validate-IsEmptyTrim($WPF_Username.Text))) {
    try {
      $UserToUnlock = $($WPF_Username.Text)
      Unlock-ADAccount -Identity $UserToUnlock

      $title = "Unlock user"
      $Message = "$($WPF_Username.Text) is unlocked"
      $Type = "Info"
    
      [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | out-null
      $path = Get-Process -id $pid | Select-Object -ExpandProperty Path
      $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
      $notify = new-object system.windows.forms.notifyicon
      $notify.icon = $icon
      $notify.visible = $true
      $notify.showballoontip(10,$Title,$Message, [system.windows.forms.tooltipicon]::$Type) 

    }
    catch {
      $title = "Unlock user"
      $Message = "$($WPF_Username.Text) is unlocked"
      $Type = "Error"
    
      [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | out-null
      $path = Get-Process -id $pid | Select-Object -ExpandProperty Path
      $icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path)
      $notify = new-object system.windows.forms.notifyicon
      $notify.icon = $icon
      $notify.visible = $true
      $notify.showballoontip(10,$Title,$Message, [system.windows.forms.tooltipicon]::$Type) 
    }
    
  }
})


$Form.ShowDialog() | Out-Null