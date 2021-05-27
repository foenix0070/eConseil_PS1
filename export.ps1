#Add-PSSnapin Microsoft.SharePoint.PowerShell
#[System.Reflection.Assembly]::Load("Microsoft.Office.Server")
Add-PSSnapin Microsoft.SharePoint
Add-PSSnapin Microsoft.SharePoint.UserProfile
[System.Reflection.Assembly]::Load("Microsoft.Office.Server")
[System.Reflection.Assembly]::Load("Microsoft.SharePoint") #,Version=14.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c"#
[System.Reflection.Assembly]::Load("Microsoft.SharePoint.UserProfile") #,Version=14.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c
clear-host

#CONSTANTES
$const_Conseil_Prefix = 'eConseil_';
$const_Calendar_WorkSpace_Field = 'Workspace';
$const_Information = 'Information';
$const_Exception = 'Exception';
$const_File = 'File';

#Variables
$var_BackUpFolder = 'D:\\';
$var_SiteUrl = "http://econseilv1migration.gouv.ci";
$var_ListPrincipalName = "Calendrier";
$var_StartDate = '';
$var_EndDate = '';


#function helpers

function Create-Folder{
  param ( [string] $p_folderName )
   if ( !(Test-Path -Path $p_folderName)) {
    New-Item -Path  $p_folderName -ItemType Directory -Force | Out-Null;
  }
}

Function Write-Log {
  [CmdletBinding()]
  Param(
    [Parameter(Mandatory = $False)]
    [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
    [String] $Level = "INFO",
    [Parameter(Mandatory = $True)]
    [string] $Message,
    [Parameter(Mandatory = $False)]
    [string]$logfile)

  $Stamp = (Get-Date).toString("dd/MM/yyyy HH:mm:ss")
  $Line = "$Stamp  $Level  $Message"
  If ($logfile) {
    Add-Content $logfile -Value $Line
    Write-Output $Line
  }

}

function Write-LogMessage {
  param ($p_file_content , [Parameter(Mandatory = $False)]
    [ValidateSet("INFO", "WARN", "ERROR", "FATAL", "DEBUG")]
    [String] $p_Level = "INFO" )
  $fileName = $backUpFolder + '\export-econseil.log';
  Write-Log -Level $p_Level -Message $p_file_content -logfile $fileName ;
}

function Write-RecapInformations {
  param (
    [String] $ListeName = '',
    [String] $BackFolderPath = '',
    [string] $StartDate,
    [string] $EndDate
  )
  Clear-Host;
  Write-host  'Recapitulatiopn de la saisie'   -ForegroundColor White;
  $StartDate = (Get-Date $StartDate).toString("dd/MM/yyyy")
  $EndDate = (Get-Date $EndDate).toString("dd/MM/yyyy")
  Write-host  " * Site à sauvegarder : $var_SiteUrl" -ForegroundColor White;
  Write-host  " * Base de données des conseils : $ListeName "   -ForegroundColor White;
  Write-host  " * Chemin de sauvegarde : $BackFolderPath "   -ForegroundColor White;
  Write-host  " * Sauvegarde des données de $StartDate à $EndDate "   -ForegroundColor White;
}

function Get-Calendar_WorkSpace{
  param ($p_spItem)
  $f_url_str = "";
  if(![string]::IsNullOrEmpty( $p_spItem[$const_Calendar_WorkSpace_Field].ToString().replace("%20", ' ').Split(",")[0] )) {
     $f_url_str = $p_spItem[$const_Calendar_WorkSpace_Field].ToString().Trim().replace("%20", ' ') ;
  }
  return $f_url_str;
}

function Get-QueryString {
  param ( [string] $StartDate, [string]$EndDate)
  $startDate = Get-Date $StartDate
  $endDate = Get-Date $EndDate
  $QueryString = '<Where><And><Geq><FieldRef Name="Created" /><Value Type="DateTime" IncludeTimeValue="True">' ;
  $QueryString += [microsoft.sharepoint.utilities.sputility]::CreateISO8601DateTimeFromSystemDateTime($startDate) ;
  $QueryString += '</Value></Geq><Lt><FieldRef Name="Created" /><Value Type="DateTime" IncludeTimeValue="True">' ;
  $QueryString += [microsoft.sharepoint.utilities.sputility]::CreateISO8601DateTimeFromSystemDateTime($endDate) ;
  $QueryString += '</Value></Lt></And></Where>';
  return  $QueryString;
}

#Count file from librairy folder
Function Get-Count_Files($Folder) {
  [int] $fileCount = 0;
  foreach ($file in $Folder.Files) { $fileCount += 1; }

  foreach ($SubFolder in $Folder.SubFolders) {
    if ($SubFolder.Name -ne "Forms") {
      $k = Get-Count_Files($Subfolder);
      $fileCount = $fileCount + $k;
    }
  }
  return $fileCount ;
}

#Count file from spWeb
function Get-Count_File_From_Web ($web) {
  $fileCount = 0;
  foreach ($list in $Web.Lists) {
    if (($List.BaseType -eq "DocumentLibrary") -and ($List.Hidden -eq $false) )
    { $fileCount += Get-Count_Files($List.RootFolder) ; }
  }
  return $fileCount;
}






#Get information to show
function  Get-MeetingInformation {
  param ($backUpFolder, $spWeb_Url, $startDate , $endDate )
  $arrMeeting = @();
  $count = 0;
  $spqQuery = New-Object Microsoft.SharePoint.SPQuery;
  $spqQuery.ViewAttributes = "Scope = 'Recursive'"
  $spqQuery.RowLimit = 5000;

  $queryString = Get-QueryString -EndDate $endDate -StartDate $startDate;


  $msg = 'Requete à executer :: ' + $queryString;
  Write-LogMessage -p_Level INFO -p_file_content    $msg ;


  $oWeb = Get-SPWeb $spWeb_Url;
  $spqQuery.Query = $queryString

  $sp_list_item_col = $oWeb.Lists[$var_ListPrincipalName ].getItems($spqQuery);
  $msg = 'Calcul des éléménts à sauvegarder '
  Write-LogMessage -p_Level INFO -p_file_content    $msg ;

  $sp_list_item_col | ForEach-Object {
    try {
      $msg = ' * Debut Traitement :: ' + $_["Title"]
      Write-LogMessage -p_Level INFO -p_file_content    $msg ;
      $wsp_url = Get-Calendar_WorkSpace $_;
      $wsp_web = Get-SPWeb $wsp_url;
      $wsp_file_count = Get-Count_File_From_Web( $wsp_web) ;

      $item = New-Object PSObject;
      $item | Add-Member -MemberType NoteProperty -Name "Title" -Value $_["Title"];
      $item | Add-Member -MemberType NoteProperty -Name "WorkSpace" -Value $wsp_url;
      $item | Add-Member -MemberType NoteProperty -Name "Fichiers" -Value $wsp_file_count;

      $arrMeeting += $item;
      $count ++;

      $msg = ' * WorkSpace trouvé :: ' + $wsp_url;
      Write-LogMessage -p_Level INFO -p_file_content    $msg ;

      $msg = ' * Nombre de fichiers trouvés :: ' +  $wsp_file_count;
      Write-LogMessage -p_Level INFO -p_file_content    $msg ;

      $msg = ' * Fin Traitement :: ' + $_["Title"]
      Write-LogMessage -p_Level INFO -p_file_content    $msg ;
    }
    catch {
      Write-LogMessage -p_Level ERROR -p_file_content  $_.Exception.Message ;
    }
  }

  $initFilePath = $backUpFolder + '\calendar-init_info.log';
  $arrMeeting | Select-Object * | Format-Table | Out-String -width 255 | Out-File -Encoding utf8 $initFilePath

  $msg = "$count éléments du calendrier ont été traités";
  Write-LogMessage -p_Level INFO -p_file_content  $msg ;

  Write-Host "Les details des informations ici : $$initFilePath";

}





#Prompt Console UI
write-host 'Adresse duy site eConseil'
$var_SiteUrl = Read-Host 'Indiquer URL du site'

write-host 'Liste prinicipale de conseils'
$var_ListPrincipalName = Read-Host 'Indiquer la liste des conseils'

write-host 'Dossier de sauvegarde'
$var_BackUpFolder = Read-Host 'Indiquer le chemin du dossier de sauvegarde'

write-host 'Saisie des parametre de periode'
$var_StartDate = Read-Host 'Saisir la date de debut dd/MM/yyyy'
$var_EndDate = Read-Host 'Saisir la date de fin dd/MM/yyyy'


#Initialisation des variables
$var_StartDate = $var_StartDate;
$var_EndDate = $var_EndDate;
$var_BackUpFolder = $var_BackUpFolder;
$var_ListPrincipalName = $var_ListPrincipalName;
$var_SiteUrl = $var_SiteUrl;

#Execution du script
Write-RecapInformations -ListeName $var_ListPrincipalName -BackFolderPath $var_BackUpFolder -StartDate $var_StartDate -EndDate $var_EndDate;

$action = Read-Host "Voulez vous continuer la sauvegarde (O/N) ?"

if ( $action -eq 'O' -or $action -eq 'o') {

  Create-Folder -p_folderName $var_BackUpFolder;

  Get-MeetingInformation -spWeb_Url $var_SiteUrl -startDate $var_StartDate -endDate $var_EndDate -backUpFolder  $var_BackUpFolder ;

  $action = Read-Host "Voulez vous continuer la sauvegarde (O/N) ?"
  if ( $action -eq 'O' -or $action -eq 'o') {




  }else{
    exit;
  }
}



