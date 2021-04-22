#Add-PSSnapin Microsoft.SharePoint.PowerShell
#[System.Reflection.Assembly]::Load("Microsoft.Office.Server")
 Add-PSSnapin Microsoft.SharePoint
 Add-PSSnapin Microsoft.SharePoint.UserProfile
 [System.Reflection.Assembly]::Load("Microsoft.Office.Server")
 [System.Reflection.Assembly]::Load("Microsoft.SharePoint") #,Version=14.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c"#
 [System.Reflection.Assembly]::Load("Microsoft.SharePoint.UserProfile") #,Version=14.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c
clear-host

#CONSTANTES
$constant_Calendar_Prefix = 'Cal_';
$constant_Calendar_WorkSpace_Field = 'Workspace';
$information = 'Information';
$erreur = 'Exception';
$fichier = 'File';


#Console UI
write-host 'Saisie des parametre de periode'
$start_date = Read-Host 'Saisir la date de debut'
$end_date = Read-Host 'Saisir la date de fin'
write-host 'Dossier de sauvegarde'
$backUpFolder = Read-Host 'Indiquer le chemin du dossier de sauvegarde'

#Variables
$backUpFolder = $backUpFolder;
$siteUrl="http://econseilv1migration.gouv.ci"
$calFolder = "" ;
$listPrincipalName = "Calendrier";
$listPrincipalName = "Agenda des Conseils" ;
$startDate = Get-Date $start_date
$endDate = Get-Date $end_date

#Start Script

#Write information in log file
function Write-In-Log_File {
  param ($p_file_name,  $p_file_content , $p_message_type = 'Informations')

  $today = Get-Date;
  $Contenu = "$today ===> [$p_message_type]   $p_file_content";
  try {
     Add-Content -Path $p_file_name -Value $Contenu
  }
  catch {
      throw $_.Exception.Message
  }

}

#Create Log file or write on it
function Create-Log_File {
  param ($p_file_content ,  $p_message_type = 'Informations')
  $fileName = $backUpFolder + '\export-econseil.log';
  if (-not(Test-Path -Path $fileName -PathType Leaf)) {
      try {
          $null = New-Item -ItemType File -Path $fileName -Force -ErrorAction Stop
          Write-Host "Le fichier de log [$file] a été crée."
      }
      catch {
          throw $_.Exception.Message
      }
  }
  Write-In-Log_File -p_file_name  $fileName -p_file_content $p_file_content -p_message_type  $p_message_type;
}

#Write information in Ui et log file
function Write-Progression {
  param ( [string] $Texte, [System.ConsoleColor] $BackgroundColor , $p_message_type = 'Informations'  )
  Write-Host $Texte -ForegroundColor $BackgroundColor;
  Create-Log_File  -p_file_content  $Texte  -p_message_type  $p_message_type;
}

### SharePoint 2010 Function ###

#Treat WorkSpace Url
function Get-Calendar_WorkSpace{
  param ($p_spItem)
  $f_url_str = "";
  if(![string]::IsNullOrEmpty( $p_spItem[$constant_Calendar_WorkSpace_Field].ToString().replace("%20", ' ').Split(",")[0] )) {
     $f_url_str = $p_spItem[$constant_Calendar_WorkSpace_Field].ToString().Trim().replace("%20", ' ') ; #Split(",")[0].ToString()
  }
  return $f_url_str;
}


#Count file from librairy folder
Function Get-Count_Files($Folder)
{
  $fileCount = 0;
    foreach($file in $Folder.Files)
	  {
      $fileCount += 1;
	  }
     foreach ($SubFolder in $Folder.SubFolders)
        {
		    if($SubFolder.Name -ne "Forms")
		    {
          $fileCount  = $fileCount + Get-Count_Files($Subfolder)
			  }
		}
    return $fileCount ;
 }

 #Count file from spWeb
function Get-Count_File_From_Web ($web) {
  $fileCount = 0;
  foreach($list in $Web.Lists)
  {
    if(($List.BaseType -eq "DocumentLibrary") -and ($List.Hidden -eq $false) )
    {
      $fileCount+= Get-Count_Files($List.RootFolder)
     }
  }
  return $fileCount;
}


function  Get-MeetingInformation  { param ($startDate , $endDate )
  $arrMeeting = @()





}






#End Script










Write-Progression -Texte 'Test ' -p_message_type $information -BackgroundColor Red

Read-Host

