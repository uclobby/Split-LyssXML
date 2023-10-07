
<#PSScriptInfo

.VERSION 1.1

.GUID 37a7aee8-54d0-423f-bb40-07f2d39a739a

.AUTHOR David Paulino

.COMPANYNAME UC Lobby

.COPYRIGHT

.TAGS Lync LyncServer SkypeForBusiness SfBServer XML

.LICENSEURI

.PROJECTURI

.ICONURI

.EXTERNALMODULEDEPENDENCIES 

.REQUIREDSCRIPTS

.EXTERNALSCRIPTDEPENDENCIES

.RELEASENOTES
  Version 1.0: 2018/08/06 - Initial release.
  Version 1.1: 2023/10/07 - Updated to publish in PowerShell Gallery.

.PRIVATEDATA

#>

<# 

.DESCRIPTION 
 Split files LyssXML export files so they can be reimported to Lyss, if the files are larger the import might fail. 

#> 

[CmdletBinding()]
param(
[parameter(ValueFromPipeline = $false, ValueFromPipelineByPropertyName = $true)]
    [string] $LyssXMLFile
	)

#Number of Queues Items per File
$ItemPerFile = 1500

if($LyssXMLFile) {

    if(Test-Path -path $LyssXMLFile) {

	    [xml]$XmlDocument = Get-Content -Path $LyssXMLFile

	    $xmlcontent = "<LyssQueueItem xmlns=""http://schemas.microsoft.com/RtcServer/2012/11/lyssimpexp"" Version=""1""><QueueItems></QueueItems></LyssQueueItem>"
	    $doc = New-Object -TypeName System.Xml.XmlDocument
	    $doc.LoadXml($xmlcontent)

	    $itemnodes = $XmlDocument.LyssQueueItem.QueueItems.ChildNodes
	    $itemCount =  $itemnodes.Count
	    $SplitedItems = 0

	    $i = 1 
	    $fileInd = 0
	    $fileNameBase = $LyssXMLFile.Substring(0,$LyssXMLFile.Length-4)
	    $SaveXML = $false

	    foreach ($item in $itemnodes){
		
		    if($i -ge $ItemPerFile){
			    $fileName  = $fileNameBase + $fileInd + ".xml"
			    $SplitedItems += $doc.LyssQueueItem.QueueItems.ChildNodes.Count
			    $doc.Save($fileName)
			    $i = 1 
			    $fileInd ++
			    $doc = New-Object -TypeName System.Xml.XmlDocument
			    $doc.LoadXml($xmlcontent)
			    $SaveXML = $false
		    } else {
			    $SaveXML = $true
		    }
		    $i++

		    [Void]$doc.LyssQueueItem.FirstChild.AppendChild($doc.ImportNode($item, $true))
	    }
	    #Make sure we save the last chunk
	    if($SaveXML) {
		    $fileName  = $fileNameBase  +$fileInd + ".xml"
		    $doc.Save($fileName)
		    $SplitedItems += $doc.LyssQueueItem.QueueItems.ChildNodes.Count
	    }

	    if($SplitedItems -eq $itemCount){
		    Write-Host "All Queued Items were splited:" $SplitedItems -ForegroundColor Green 
	    } else {
		    Write-Host "Not all items were splited:"$SplitedItems $itemCount -ForegroundColor Red
	    }

    }
}