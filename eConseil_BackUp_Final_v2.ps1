﻿#Add-PSSnapin Microsoft.SharePoint.PowerShell
#[System.Reflection.Assembly]::Load("Microsoft.Office.Server")

# clear-host
 Add-PSSnapin Microsoft.SharePoint
 #Add-PSSnapin Microsoft.SharePoint.UserProfile
# [System.Reflection.Assembly]::Load("Microsoft.Office.Server")
# [System.Reflection.Assembly]::Load("Microsoft.SharePoint") #,Version=14.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c"#
# [System.Reflection.Assembly]::Load("Microsoft.SharePoint.UserProfile") #,Version=14.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c
 #clear-host

#CONSTANTES
$constant_Calendar_Prefix = 'Cal_';
$constant_Calendar_WorkSpace_Field = 'Workspace';

#Variables
$backUpFolder = "E:\eGouv";
$siteUrl="http://econseilv1migration.gouv.ci" #"http://econseil-test.egouv.ci/";
$calFolder = "" ;
$listPrincipalName = "Calendrier";
$listPrincipalName = "Agenda des Conseils" ;
$startDate = Get-Date -Month 11 -Day 23 -Year 2000 -Hour 00 -Minute 00 -Second 00;
$endDate = Get-Date -Month 11 -Day 23 -Year 2021 -Hour 00 -Minute 00 -Second 00;

#Start Script
cd $backUpFolder;

function Get-Fields{
   param ($p_spList)
     return $p_spList.Fields | Where-Object  { ($_.Hidden -eq $false) -and ($_.ReadOnlyField -eq $false)} | select Title, InternalName 
}

function Get-Calendar_WorkSpace{
   param ($p_spItem)
   $f_url_str = "";
   if(![string]::IsNullOrEmpty( $p_spItem[$constant_Calendar_WorkSpace_Field].ToString().replace("%20", ' ').Split(",")[0] )) { 
      $f_url_str = $p_spItem[$constant_Calendar_WorkSpace_Field].ToString().Trim().replace("%20", ' ') ; #Split(",")[0].ToString()
   } 
   return $f_url_str;
}

function Create-XmlFileWithContent{
   param ( [string] $p_fileName , [string] $p_fileContent  )
    $p_fileContent = '<?xml version="1.0" encoding="utf-8"?>' + $p_fileContent;
    New-Item -Path $p_fileName  -ItemType File -Force  ;
    Set-Content -Path  $p_fileName -Value $p_fileContent -Force | Out-Null;
}

function Create-Folder{
   param ($p_folderName )
    New-Item -Path $p_folderName  -ItemType Directory -Force | Out-Null;
}

function Create-ListItemProperty{  
 param ( $p_parent , $p_item , $p_spField)
 $str = ""
  $p_parent = [String] $p_parent;
   
 if($p_item -ne $null){
  foreach ($t_prop in $p_spField ) {
  
  
    $str += '<item ' ;
    $str += 'InternalName="' + ($t_prop.InternalName) + '" ';
    $str += 'Title="' + ($t_prop.Title) + '" ';
    $str += 'Value="' + ($p_item[$t_prop.InternalName]) + '" ';
    $str += '/>' ;
  }
  }
  $str = "<" + $p_parent + ">" + $str +  "</"+ $p_parent +">";
  return $str;
}

function WriteFileProperty {
  param(  $p_file , $p_destination)
    
   $str = '';
   $p_destination = $p_destination +'/' + $p_file.Name + '.xml';
   $f_item = $p_file.Item; 
   $f_Filefieds = Get-Fields $f_item.ParentList;
 
   $str += Create-ListItemProperty -p_parent "DocItem" -p_item $f_item -p_spField $f_Filefieds; 
    
   Create-XmlFileWithContent -p_fileName $p_destination -p_fileContent $str; 
}


function ProcessFolder {
    param($p_web, $folderUrl , $destination )
    
    $f_sw = Get-SPWeb $p_web;
    
    $folder = $f_sw.GetFolder($folderUrl)
    foreach ($file in $folder.Files) {
        #Ensure destination directory
        $destinationfolder = $destination + "/" + $folder.Url 
        if (!(Test-Path -path $destinationfolder))
        {
            $dest = New-Item $destinationfolder -type directory 
        }
        #Download file
        $binary = $file.OpenBinary()
        $stream = New-Object System.IO.FileStream($destinationfolder + "/" + $file.Name), Create
        $writer = New-Object System.IO.BinaryWriter($stream)
        $writer.write($binary)
        $writer.Close()
    }
        
        
    
}




function ProcessLibrairyFolderDownload {
    param( $p_web_url, $p_folderUrl , $p_destination) 
    $p_web =  Get-SPWeb $p_web_url
   
    $f_folder = $p_web.GetFolder($p_folderUrl)
      
    
    foreach ($file in $f_folder.Files) {     
        $f_destinationfolder = $p_destination ;
        $f_destinationfolder
        if (!(Test-Path -path $f_destinationfolder))
        {
            Create-Folder $f_destinationfolder 
        }
        write-host $file.Name
        $binary = $file.OpenBinary()
        $stream = New-Object System.IO.FileStream($f_destinationfolder + "/" + $file.Name), Create
        $writer = New-Object System.IO.BinaryWriter($stream)
        $writer.write($binary)
       $writer.Close()
        WriteFileProperty  -p_file $file -p_destination $f_destinationfolder;   
    }
    
    
}

function Get-LirairiesFiles{
 param ($p_item  )

 $str_url = Get-Calendar_WorkSpace -p_spItem $p_item; 
 $f_sw = Get-SPWeb $str_url;
 $f_lists = $f_sw.lists  | Where-Object  { ($_.hidden -eq $false) -and ($_.IsSiteAssetsLibrary -eq $false) -and ($_.BaseType -eq "DocumentLibrary")} ;
  foreach ($l  in $f_lists ) {
   
       $str_folder_name = $l.Rootfolder.Name; 
       Create-Folder $str_folder_name;

       $strCurrentDirectory =  Get-Location  ;
       $str_folder_name =   $strCurrentDirectory.Path + "/" + $str_folder_name ;
       
       $DocLibItems = $l.Items
       
       foreach ($DocLibItem in $DocLibItems) {
        if($DocLibItem.Url -Like "*.pdf") {
             
            $File = $f_sw.GetFile($DocLibItem.Url)
            $Binary = $File.OpenBinary()
            $Stream = New-Object System.IO.FileStream( $str_folder_name + "\" + $File.Name), Create
            $Writer = New-Object System.IO.BinaryWriter($Stream)
            $Writer.write($Binary)
            $Writer.Close()
            
             WriteFileProperty  -p_file $File -p_destination  $str_folder_name ; 
        }
    }
       
     
  }
  
  return $str;
}
  
function Create-ListProperty{
 param ($p_item, $p_str  )

  $str = $p_str;
  $f_url = Get-Calendar_WorkSpace  $p_item;
  $f_sw = Get-SPWeb  $f_url; 


 $Lists = $f_sw.lists  | Where-Object  { ($_.hidden -eq $false) } #-and ($_.BaseType -eq "GenericList") -and ($_.IsSiteAssetsLibrary -eq $false) 
  foreach ($l  in $Lists ) {
   
    $str += '<propertyList ' ;
    $str += 'RootFolder="' + ($l.RootFolder.Name) + '" ';
    $str += 'Title="' + ($l.Title) + '" ';
    $str += 'Description="' + ($l.Description) + '" ';
    $str += '>' ;
     
     $lItems = $l.GetItems();
     Get-Fields $l.Fields 
 
       foreach ($it in $lItems ) { 
          $str +=  Create-ListItemProperty -p_parent "ListItem" -p_item $it -p_spField $SPLField; 
       } 
   
     $str +='</propertyList>';
    
     
  }
  
  return $str;
}

function Create-FullProperty{
 param ($p_item , $p_fields , $p_str)
   
   $str +=  Create-ListItemProperty -p_parent "ListItem" -p_item $p_item  -p_spField $p_fields;
   $str+= Create-ListProperty   -p_item $p_item -p_str $p_str;
 
   return $str;
 }
  
function Create-CalendarFolder{
 param ($p_items , $p_calendarFields )
  $strName = "";
  $str = '';
  foreach ($item in $p_items ) {
      $str = '';
      $strName  = $constant_Calendar_Prefix  + $item.ID  ;
      Create-Folder $strName;
      cd $strName;
       
     $str = Create-FullProperty -p_item $item -p_fields $p_calendarFields -p_str $str ; 
     $strName= $strName +'.xml' ;
     Create-XmlFileWithContent -p_fileName $strName -p_fileContent $str;
    
     Get-LirairiesFiles $item;
    
    cd ..;
  }
}

#Recuperation des éléments du calendrier
#[microsoft.sharepoint.utilities.sputility]::CreateISO8601DateTimeFromSystemDateTime
#Clear-Host
$QueryString = '<Where><And><Geq><FieldRef Name="Created" /><Value Type="DateTime" IncludeTimeValue="True">' ;
$QueryString += [microsoft.sharepoint.utilities.sputility]::CreateISO8601DateTimeFromSystemDateTime($startDate) ;
$QueryString +=  '</Value></Geq><Lt><FieldRef Name="Created" /><Value Type="DateTime" IncludeTimeValue="True">' ;
$QueryString += [microsoft.sharepoint.utilities.sputility]::CreateISO8601DateTimeFromSystemDateTime($endDate) ;
$QueryString += '</Value></Lt></And></Where>';
 

$spqQuery = New-Object Microsoft.SharePoint.SPQuery
$spqQuery.ViewAttributes = "Scope = 'Recursive'"
$spqQuery.RowLimit = $SPList.ItemCount
$spqQuery.Query =   $QueryString
 
 
$oWeb=Get-SPWeb $siteUrl;


$SPCalendar=$oWeb.Lists[$listPrincipalName];
$SPCalendarItems =  $SPCalendar.getItems($spqQuery);
$SPCalendarField =  Get-Fields $SPCalendar;

#Execution de la sauvegarde des éléments du calendrier et de l'espace dédié
Clear-Host
write-host 'Debut de la sauvegarde ' -foregroundcolor DarkGreen -backgroundcolor white
Create-CalendarFolder -p_items $SPCalendarItems -p_calendarFields $SPCalendarField ;
write-host 'Fin de la sauvegarde ' -foregroundcolor DarkGreen -backgroundcolor white