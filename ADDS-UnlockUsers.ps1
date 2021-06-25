#========================================================================
#
# Tool Name	: ADDS-UnlockUsers
# Author 	: METRAL Emile 
# Gitlab / Github : EBMBA
#
#========================================================================

[Void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework') 
foreach ($item in $(gci .\assembly\ -Filter *.dll).name) {
    [Void][System.Reflection.Assembly]::LoadFrom("assembly\$item")
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

$Form.Add_ContentRendered({
  $WPF_Validate.IsEnabled = $false
})

foreach ($UserLock in $UsersLocks) {
  $GroupsList = New-Object PSObject
  $GroupsList = $GroupsList | Add-Member NoteProperty SamAccountName $UserLock.Login -passthru
  $GroupsList = $GroupsList | Add-Member NoteProperty Name $UserLock.Name -passthru	
  $WPF_Users.Items.Add($GroupsList) > $null
}

$WPF_Username.Add_TextChanged({
  if(Get-ADUser -Identity $($WPF_Username.Text)){
    $WPF_Username.Background = [System.Windows.Media.Brushes]::PaleGreen
    $WPF_Validate.IsEnabled = $True
  }
  else {
    $WPF_Username.Background = [System.Windows.Media.Brushes]::PaleVioletRed
    $WPF_Validate = $false
  }
})

$WPF_Validate.Add_Click({
  if((Validate-IsEmptyTrim($WPF_Username)) -and ($($WPF_Users.SelectedItems).Length -gt 0)){
    $UsersToUnlock = $($WPF_Users.SelectedItems).Login
    foreach($UserToUnlock in $UsersToLock){
      Unlock-ADAccount -Identity $UserToUnlock
    }
  }
  elseif (-not(Validate-IsEmptyTrim($WPF_Username)) -and ($($WPF_Users.SelectedItems).Length -le 0)) {
    $UserToUnlock = $($WPF_Username.Text)
    Unlock-ADAccount -Identity $UserToUnlock
  }
})


$Form.ShowDialog() | Out-Null