Add-type -assembly “Microsoft.Office.Interop.Outlook” | out-null
$olFolders = “Microsoft.Office.Interop.Outlook.olDefaultFolders” -as [type] 
$o = New-Object -comobject outlook.application
$n = $o.GetNamespace("MAPI")
$f = $n.GetDefaultFolder($olFolders::olFolderInbox).Folders("Sample Folder") 
$index = 0
[String[]]$Array = @()
[String[]]$SortedArray = @()
$TempFile = "Sample Directory"
$DestinationFolder = "Sample Directory"

  $f.Items| foreach {
     $SendName = $_.SenderName
     $_.attachments| foreach {
     $object = $_.filename
        If ($object.Contains("String") )
        {
        $Array += $object
        $index += 1
        }         
     }    
   }
 
    $SortedArray = $Array | Sort-Object
   
  
    

    #Never Comment Anything above this line :)

    #Normal Usage Section - Leave Catch-Up section commmented

    $Length = $SortedArray.Length
    $ReportIndex = $Lenght - 1
    $ReportFilename = $SortedArray[$ReportIndex]
    
    
    $f.Items| foreach {
     $SendName = $_.SenderName
     $_.attachments| foreach {
     $object = $_.filename 
    
        If ($object.Contains($ReportFilename) )
        {
        $_.saveasfile(($TempFile))
        }
        
      }
        
    }
    
    Expand-Archive -LiteralPath $TempFile -DestinationPath $DestinationFolder
    Remove-Item –path $TempFile 



    #Catch up Section - If catching up, comment whole Normal Usage section.
    #The way the catch-up section works is that basically you make a Temporary array with all the reports that you've missed and the script will go and save every report on that array.
    #Comment this section when done

    #$SortedArray.Length
    #$TempArray = $SortedArray[]
    #Write-Host $TempArray
    #
    #$f.Items| foreach {
    # $SendName = $_.SenderName
    # $_.attachments| foreach {
    # $object = $_.filename 
    #
    #    for( $n=0;$n -le $TempArray.Length - 2 ;$n++) {
    #        If ($object.Contains($TempArray[$n]) )
    #        {
    #        $_.saveasfile(($TempFile))
    #        Expand-Archive -LiteralPath $TempFile -DestinationPath $DestinationFolder
    #        Remove-Item –path $TempFile 
    #        {break}
    #
    #        }
    #    
    #    }
    #    }             
    #  }

    #closes the outlook process if it was not running before executing this script (keeps things squeaky clean).

    if ($outlookWasAlreadyRunning -eq $false)
{
    Get-Process "*outlook*" | Stop-Process –force
}
        