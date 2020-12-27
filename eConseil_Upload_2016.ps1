Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Runtime.dll" 
Add-Type -Path "c:\Program Files\Common Files\microsoft shared\Web Server Extensions\15\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll" 

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

#Constantes
$const2016CalendarListName = 'Calendrier';
$const2016GestionTextDocsName = 'GestionTexte';
$const2016ConseilDocsListName = 'ConseilDocs';

$constArrayFieldsToIgnore = 'ContentType', 'Attachments', 'ParticipantsPicker' , 'FreeBusy', 'Facilities';
$constArrayGestTtextFields =  'TypeTexte' , 'Ministere', 'Statut',  'RoleNumber', 'RoleDate', 'JONumber' , 'JODate', 'TypeConseil', 'AnalystGroups', 'SignatureDate', 'TransmissionDate', 'SentDate', 'AnalyseDate', 'EnrolSendDate', 'PublishSendDate', 'MinistereID', 'TypeTexteID', 'CommentaireDepot', 'History', 'AReserve', 'OriginalName';
$constArrayDateFields = 'EndDate', 'EventDate', 'RoleDate', 'JODate', 'SignatureDate', 'TransmissionDate' , 'SentDate', 'AnalyseDate', 'EnrolSendDate'   ;

 
 function Write-Progression {
    param ( [string] $Texte) 
    Write-Host $Texte -ForegroundColor Green
 }


function is-InArray{
  param ($Array, [string] $Value) 

  $eleme = [Array] $Array; 
  foreach ($u in $eleme) { 
    foreach ($el in $u) {  
      if ($el.trim() -eq $Value.trim()) { return $true; } 
    }
  }
  return $false;
}

function Get-FildValue {
  param ( $FiedName , $Value)

  $v = is-InArray -Array $constArrayDateFields -Value FiedName ;
  if ($v -eq $true) {   
    return  [Datetime]::ParseExact($Value, 'MM/dd/yyyy HH:mm:ss', $null)   
  }
  return $Value;
}

function Repair-XmlString {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$inXML
  ) 
  $rPattern = "[^\x09\x0A\x0D\x20-\xD7FF\xE000-\xFFFD\x10000\x10FFFF]"; 
  return [System.Text.RegularExpressions.Regex]::Replace($inXML, $rPattern, "");
}

function Get-ClientContext {
  param ( $p_url, $p_login, $p_pwdstring  )
  $pwd = ConvertTo-SecureString $p_pwdstring -AsPlainText -Force;
  $ctx = New-Object Microsoft.SharePoint.Client.ClientContext($p_url);
  [Net.ServicePointManager]::SecurityProtocol = "Ssl3", "Tls", "Tls11", "Tls12";  
  [System.Net.ServicePointManager]::ServerCertificateValidationCallback = { $true };  
  $credentiale = New-Object System.Net.NetworkCredential($p_login, $pwd) ;
  $ctx.Credentials = $credentiale;
  return $ctx;
}

function Add-CalendarListItem {
  param ([Microsoft.SharePoint.Client.ClientContext] $ctx, [string] $listName, $listItemToAdd)

  try {  
    $lists = $ctx.web.Lists;
    $list = $lists.GetByTitle($listName); 
         
    $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation;
    $listItem = $list.AddItem($listItemInfo);
    $listItem['Statut'] = 'Cloture';
     
    $Ligne = $listItemToAdd.Item | Select InternalName , Value;
    foreach ($c in $Ligne) { 
      $v = is-InArray -Array $constArrayFieldsToIgnore -Value $c.InternalName ;
      if ($v -eq $false) {  
         $listItem[$c.InternalName]   = Get-FildValue -Value $c.Value -FiedName $c.InternalName;
      }
    }
    $listItem.Update();
    $ctx.load($lists);    
    $ctx.load($list);    
    $ctx.executeQuery();  

  
   return $listItem.Id  ;
  }  
  catch {  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
  }   
}

function Add-DocumentProperty {
  param ([Microsoft.SharePoint.Client.ClientContext] $ctx, [string] $listName, $item , $serverRelativeUrl , $CalendarEventID , $DocuementID)

  
  try {  
     
    $lists = $ctx.web.Lists;
    $list = $lists.GetByTitle($listName);  
    $listItemInfo = New-Object Microsoft.SharePoint.Client.ListItemCreationInformation;
    $listItem = $list.AddItem($listItemInfo);

    $listItem['DocLink'] = $serverRelativeUrl;
    $listItem['ConseilID'] = $CalendarEventID;
    $listItem['RubriqueID'] = '0';
    $listItem['RoleNumber'] = '0';
    $listItem['Decision'] = '';
    $listItem['DocObjet'] = '';
    $listItem['DocumentID'] =  $DocuementID ;
  
    $listItem['TypeTexteID'] = $item['TypeTexteID'];
    $listItem['MinistereID'] = $item['MinistereID'];
     
    $listItem['TypeTexte'] = $item['TypeTexte'];
    $listItem['Ministere'] = $item['Ministere'];
    $listItem['Title'] = $item['OriginalName'];
    $listItem['DocName'] = $item['OriginalName'];
    $listItem["ConseilDate"] = [System.DateTime]::Now
     
    $listItem.Update();
    $ctx.load($lists);    
    $ctx.load($list);    
    $ctx.executeQuery();  

    
  }  
  catch {  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
  }   
}

function Add-DocumentToLibrairy (){
  param ([Microsoft.SharePoint.Client.ClientContext] $ctx, $CalendarEventID , [string] $listName, $libFolder) 

  try{
  $web = $ctx.Web;
  $docLib = $web.Lists.GetByTitle($listName);
  $ctx.Load($web);
  $ctx.Load($docLib);
  $ctx.ExecuteQuery();

      $Files = Get-ChildItem -Path $libFolder.FullName | ? {$_.psIsContainer -eq $False -and $_.Extension -ne '.xml'}
      
       Foreach ($File in $Files) { 
       

        $FileFullName = $File.FullName
        $FileStream = New-Object IO.FileStream($FileFullName, [System.IO.FileMode]::Open)
        $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
        $FileCreationInfo.Overwrite = $true
        $FileCreationInfo.ContentStream = $FileStream
        $FileCreationInfo.URL = $File.Name
        $FileUpload = $docLib.RootFolder.Files.Add($FileCreationInfo )
       
        $listItem =  $FileUpload.ListItemAllFields
        $listItem["RoleDate"] = [System.DateTime]::Now
        $listItem["JODate"] = [System.DateTime]::Now
        $listItem["SignatureDate"] = [System.DateTime]::Now
        $listItem["TransmissionDate"] = [System.DateTime]::Now
        $listItem["SentDate"] = [System.DateTime]::Now
        $listItem["AnalyseDate"] = [System.DateTime]::Now
        $listItem["EnrolSendDate"] = [System.DateTime]::Now
      
        $listItem["TypeTexte"] = "Docuements"
        $listItem["Ministere"] = "Ministere"
        $listItem["MinistereID"] = "0"  
        $listItem["TypeTexteID"] = "0"  
        $listItem["CommentaireDepot"] = ""  
        $listItem["TypeTexteID"] = "0"  
        $listItem["OriginalName"] =$File.Name
        $listItem["Title"] = "Document";
        $listItem['Statut'] = 'Cloture';
        
        $listItem.Update();

         $ctx.Load($listItem)
        $ctx.Load($FileUpload)
        $ctx.ExecuteQuery()

        $id = $listItem.Id;
        

         
        Add-DocumentProperty -ctx $ctx -listName $const2016ConseilDocsListName -item $listItem -serverRelativeUrl $FileUpload.ServerRelativeUrl -CalendarEventID $CalendarEventID -DocuementID $id  ;
          
      }
  }
  catch {  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
  }   
}


function Process-RestoreBackUp {
  param ( [Microsoft.SharePoint.Client.ClientContext] $ctx , $folder )
   
  $folderType = $folder.GetType() ; 
  if ($folderType.FullName.EndsWith("DirectoryInfo")) {  
    $idListItem =0;
    Write-Progression -Texte "restauration de  $folder ..."

    $xmlFilePath = $folder.FullName + '\' + $folder.Name + '.xml';
    [xml] $xmlCalendarPropertiesFile = Repair-XmlString (Get-Content $xmlFilePath -Raw) ;
    $idListItem = Add-CalendarListItem -ctx $ctx -listName $const2016CalendarListName -listItemToAdd $xmlCalendarPropertiesFile.Document.ListItem;
    $librairiesfolder = Get-ChildItem $folder.FullName;

     foreach ($f in $librairiesfolder) { 
       $folderType = $f.GetType() ; 
         if ($folderType.FullName.EndsWith("DirectoryInfo")) {  
          Add-DocumentToLibrairy -ctx $ctx -listName $const2016GestionTextDocsName -libFolder $f -CalendarEventID $idListItem;
         }
     }
  }
}



#Assignation des variables
Write-Progression -Texte "Initialisation des variables de connexion..."
$url = "https://test-econseil2016.gouv.ci/";
$login = "GOUV\inova.econseil";
$pwdstring = "Inov@2017";
$backUpFolder = "D:\Save";

#Execution du telechargement
 Write-Progression -Texte "Debut la de connexion..."
 $clientContext = Get-ClientContext -p_url $url -p_login $login -p_pwdstring $pwdstring;

 Write-Progression -Texte "Exploration du dosier de sauvegarde..."
 $listfolder = Get-ChildItem $backUpFolder;

  Write-Progression -Texte "Debut de la restauration..."
foreach ($f in $listfolder) { 
  Process-RestoreBackUp -ctx $clientContext -folder $f;
}
 
 Write-Progression -Texte "Fin de la restauration..."