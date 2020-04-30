
$Error.Clear()
Function CopyVeryLongPath([String]$Source,[string]$PathToCreate,[string]$DocumentName,[Switch]$SourceIsUNC,[Switch]$destinationIsUNC,[Switch]$IsFolder,$Destination='Destination Path'){
    $BaseSource='\\?\'
    $BaseDestination='\\?\'
    #"SourceIsUNC:"+$SourceIsUNC
    if($SourceIsUNC){
        $BaseSource='\\?\UNC\'
        $Source=$Source.Replace("\\","")                                         
    }
    if($destinationIsUNC){
        $BaseDestination='\\?\UNC\'
        $Destination=$Destination.Replace("\\","")
    }
        $Source=$BaseSource+$Source#+'\'+$FolderName         
        $Destination=$BaseDestination+$Destination
        Write-Host "Source is: " $Source
        Write-Host "Destination is: " $Destination
  
    if($IsFolder){               
        Copy-Item -LiteralPath $Source -Destination $Destination -Force -Recurse -Verbose #| Out-Null
    }
    else{
        
        
        $Error.Clear()
        Copy-Item -LiteralPath $Source -Destination $Destination -Force -Verbose 
        #error will occured because the copy function will not find the parent file if it doesn't exist
        if($Error){
            New-Item -Path $PathToCreate -Name $DocumentName -ItemType "Directory"
            Copy-Item -LiteralPath $Source -Destination $Destination -Force -Verbose | Out-File "$Destination\errors\$($DocumentName).txt" -Append
        $Error.Clear()
        }
            
    }
    
} #end of CopyVeryLongPath

function Find-Folders ([string]$Description,[switch]$IsTarget){


    [Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms") | Out-Null
    [System.Windows.Forms.Application]::EnableVisualStyles()
    $browse = New-Object System.Windows.Forms.FolderBrowserDialog
   # $browse.SelectedPath = "\\unc"
    
    $browse.Description = $Description
    if($IsTarget){
        $browse.ShowNewFolderButton = $true
    }
    else{
        $browse.ShowNewFolderButton = $false
    }
    

    $loop = $true
    while($loop)
    {
        if ($browse.ShowDialog() -eq "OK")
        {
        $loop = $false
		
		#Insert your script here
		
        } else
        {
            $res = [System.Windows.Forms.MessageBox]::Show("You clicked Cancel. Would you like to try again or exit?", "Select a location", [System.Windows.Forms.MessageBoxButtons]::RetryCancel)
            if($res -eq "Cancel")
            {
                #Ends script
                return
            }
        }
    }
    $browse.SelectedPath
    $browse.Dispose()
    
} #end of Find-Folders

Function Edit-Excel([switch]$Replace){
    [string]$Description=$null
    if($Replace){
        $Description="and edit them"
    }
    $root=Find-Folders -Description "Select directory to open the workbooks which contain hyperlinks $Description"
    if(-not $Replace){
        $destination=Find-Folders -Description "Select directory to copy the files and folders from hyperlinks which are in workbooks" -IsTarget
    }
    $destinationPathInfo=[System.Uri]$destination
    if($destinationPathInfo.IsUnc){
    #if($hyperlink -like "*\\*"){
        $destinationIsUNC=$true
        }
    else{
        $destinationIsUNC=$false
        }
    #Set allowed file extensions
    $excelWorkbooks = Get-Childitem -Path $root -Include *.xls,*.xlsx, *.xlsm -Recurse
    #Open Excel ComObject
    $excel = New-Object -comobject Excel.Application

    #Show Excel during the process
    $excel.visible = $false
    if($Replace){
        $StringToBeReplaced=Read-Host "Write the string of path to be replaced"
        $NewString=Read-Host "Write new string"
        foreach($excelWorkbook in $excelWorkbooks){
         $workbook = $excel.Workbooks.Open($excelWorkbook)
         $workbookBaseName=[io.path]::GetFileNameWithoutExtension($excelWorkbook)
         "There are $($workbook.Sheets.count) sheets in $excelWorkbook"
          For($i = 1 ; $i -le $workbook.Sheets.count ; $i++){
             $worksheet = $workbook.sheets.item($i)
             $hyperlinks = $Worksheet.Hyperlinks
             foreach ($hyperlink in $hyperlinks){
                “Displayed text: ” + $hyperlink.TextToDisplay                
                “Hyperlink: ” + $hyperlink.Address
                $addressChange= $hyperlink.Address.Replace($StringToBeReplaced,$NewString)
                $hyperlink.Address=$addressChange
             } #end of foreach ($hyperlink in $hyperlinks)
          } #end of For($i = 1 ; $i -le $workbook.Sheets.count ; $i++)
        }#end of  foreach($excelWorkbook in $excelWorkbooks)
        $workbook.save()
        #if (!$Workbook.saved) { $workbook.save() }
        $workbook.close()
    }#end of if($Replace)
    else{
        foreach($excelWorkbook in $excelWorkbooks){
         $workbook = $excel.Workbooks.Open($excelWorkbook)
         $workbookBaseName=[io.path]::GetFileNameWithoutExtension($excelWorkbook)
         "There are $($workbook.Sheets.count) sheets in $excelWorkbook"
         Push-Location "$excelWorkbook\.." #change to parent directory of opened document or excel workbook
         [string]$InitialexcelWorkbookPath=$null
         $InitialexcelWorkbookPath=$excelWorkbook.PSParentPath ###ekei pou einai ta peiramata
         $InitialexcelWorkbookPath=$InitialexcelWorkbookPath.Replace("Microsoft.PowerShell.Core\FileSystem::","")
         For($i = 1 ; $i -le $workbook.Sheets.count ; $i++){
          $worksheet = $workbook.sheets.item($i)
          $hyperlinks = $Worksheet.Hyperlinks


          foreach ($hyperlink in $hyperlinks){
            if(($hyperlink.Address -like "*http*") -or ($hyperlink.Address -like "*www*")){
                continue
            }
            #$pathArray=$hyperlink.Address -split '\\' #Actually it is one '\' because it is a escape character
            #$LastFileOrFolderOfPath=$pathArray[$pathArray.length-1] #pass the value of last folder or file of the hyperlink
            #"Last file or directory: "+$LastFileOrFolderOfPath
            $Address=$hyperlink.Address
            $IsFolder=$false
            $IsFile=$true
            if(Test-Path $hyperlink.Address -PathType Leaf){
                $IsFolder=$false
                # $FileName=(Get-Item -Path $hyperlink.Address).BaseName
                #$extension=(Get-Item -Path $hyperlink.Address).Extension
                $FileName=[System.IO.Path]::GetFileNameWithoutExtension("$($hyperlink.Address)")
                $extension=[System.IO.Path]::GetExtension("$($hyperlink.Address)")
                $file=[System.IO.Path]::GetFileName("$($hyperlink.Address)")
                $Address="$Address\.."                        
                            
            }#if(Test-Path $hyperlink.Address -PathType Leaf)
            elseif(Test-Path $hyperlink.Address -PathType Container){
                $IsFolder=$true
                $IsFile=$false
                <#$FoldersArray=$Address.Split('\').Split('/')
                $lastFile=$FoldersArray[$FoldersArray.Length-1]
                if($lastFile -like "*.*"){
                $Address="$Address\.."
                }#>

            }

            $PathInfo=[System.Uri]$hyperlink.Address;
            if($PathInfo.IsUnc){
            #if($hyperlink -like "*\\*"){
                $SourceIsUNC=$true
            }
            else{
                $SourceIsUNC=$false
            }
            [int]$pops=0
            if(Test-Path $Address){
                Push-Location $Address
                $pops++            
            }
            elseif(Test-Path $InitialexcelWorkbookPath){
                Push-Location $InitialexcelWorkbookPath
                $pops++
                if(Test-Path $Address){
                    Push-Location $Address
                    $pops++
                }
                else{
                    $Address=$Address.Replace("..\","").Replace("../","")
                    if($IsFile){
                        $Address="$Address\.."
                    }
                    $Address="$InitialexcelWorkbookPath\$Address"
                    $Error.Clear()
                    Push-Location $Address | Out-Null
                    $pops++
                    if($Error){
                        $pops--         
                        foreach($ErrorItem in $Error){
                            $ErrorItem.Exception  | Out-File "$destination\errors\$($workbookBaseName).txt" -Append
                        }                         
                        $Error.Clear()
                        for($pop=0;$pop -lt $pops;$pop++ ){
                            Pop-Location
                        }
                       "all wrong"
                        continue #bypass the hyperlink
                    }#if($Error)
                }
            }#elseif(Test-Path $InitialexcelWorkbookPath){



                  foreach($ErrorItem in $Error){
                       $ErrorItem  | Out-File "$destination\errors\$($workbookBaseName).txt" -Append
                       }
                  $Error.Clear() 
                  Push-Location -StackName stack1

                  [string]$source=$null
                  $source=Get-Location -StackName stack1
                  $source=$source.replace("Microsoft.PowerShell.Core\FileSystem::","")
                  if(-not $IsFolder){
                    $source="$source\$file"
                  }
                  "Source before copy function :$source"
                  Pop-Location -StackName stack1
                  for($pop=0;$pop -lt $pops;$pop++ ){
                     Pop-Location
                  }
                 # Pop-Location
                 # if($SecondPop){
                  #  Pop-Location
                  #}


                  if(test-path $source -PathType Container){
                    $IsFolder=$true
                    $folder=Split-Path -Path $source -Leaf -Resolve 
                  }
                  elseif(Test-Path $source -PathType Leaf){
                    $IsFolder=$false
                    $LastDirectory=Split-Path -Path $source -Leaf -Resolve 
                    #$file=(Get-Item -Path $destination3).BaseName
                    #$extension=(Get-Item -Path $destination3).Extension
                    

                  }
                  
                  if($IsFolder){
                    [string]$destination2=$destination+'\'+$workbookBaseName+'\'+$folder
                    $i=0
                    while(Test-Path $destination2){
                        $destination2=$destination+'\'+$workbookBaseName+'\'+$folder+'_'+$i
                        $i++
                    }
                    $i=0
                    }#if($IsFolder){
                  else{
                    [string]$destination2=$destination+'\'+$workbookBaseName+'\'+$FileName+$extension
                    $i=0
                    while(Test-Path $destination2){
                        $destination2=$destination+'\'+$workbookBaseName+'\'+$FileName+'_'+$i+$extension
                        $i++
                    }
                    $i=0
                  }

                  #########################na allakso to Document.BaseName


                  CopyVeryLongPath -Source $source -Destination $destination2 -IsFolder:$IsFolder -PathToCreate $destination -DocumentName $workbookBaseName -SourceIsUNC:$SourceIsUNC -destinationIsUNC:$destinationIsUNC

                  CopyVeryLongPath -Source $source -Destination $destination2 -IsFolder:$IsFolder -PathToCreate $destination -DocumentName $workbookBaseName -SourceIsUNC:$SourceIsUNC -destinationIsUNC:$destinationIsUNC

            <#if(($IsFile)){ #checks if a hyperlink is a file
                
                
                CopyVeryLongPath -Source $hyperlink.Address -FolderName $LastFileOrFolderOfPath -Destination $destination2 -SourceIsUNC:$SourceIsUNC
            }
            else {
                
                
                CopyVeryLongPath -Source $hyperlink.Address -FolderName $LastFileOrFolderOfPath -Destination $destination2 -IsFolder -SourceIsUNC:$SourceIsUNC #-FolderName $LastFileOrFolderOfPath
            }#>
                 
        $hyperlink.Address=$destination2
        } #end foreach ($hyperlink in $hyperlinks)
     Pop-Location

     ## Save
     $workbook.save()
        #if (!$Workbook.saved) { $workbook.save() }
     $workbook.close()

    } #end For($i = 1 ; $i -le $workbook.Sheets.count ; $i++)
    }#end of foreach($excelWorkbook in $excelWorkbooks)
     
}# end of else
     $excel.quit()
     $excel = $null
     $Description=$null
     $Replace=$false
     [gc]::collect()
     [gc]::WaitForPendingFinalizers()
}# end of Function Edit-Excel([switch]$Replace)


Function Edit-Word([switch]$Both,[switch]$Replace){
    $passwd="" #password for opening docs
    $passwdTemplate=""
    $passwdWrite="" #password for editing read-only protected docs
    [string]$Description=$null
    if($Replace){
        $Description="and edit them"
    }
    $root=Find-Folders -Description "Select directory to open the documents which contain hyperlinks $Description"

    if(-not $Replace){
        $destination=Find-Folders -Description "Select directory to copy the files and folders from hyperlinks which are in the documents" -IsTarget
    }
    $destinationPathInfo=[System.Uri]$destination
    if($destinationPathInfo.IsUnc){
    #if($hyperlink -like "*\\*"){
        $destinationIsUNC=$true
        }
    else{
        $destinationIsUNC=$false
        }
    $Documents = Get-ChildItem -Path $root -Filter "*.doc*" -Recurse

    $Word = New-Object -comobject word.application
    $Word.Visible = $false
     if($Replace){
        [string]$StringToBeReplaced=Read-Host "Write the string of path to search"
        [string]$NewString=Read-Host "Write new string"
        foreach ($Document in $Documents){
            $DocumentToEdit = $Word.documents.open($Document.FullName,$null, $false, $null, $passwd, $passwdTemplat, $false, $passwdWritee)
            "Processing file: {0}" -f $DocumentToEdit.FullName    
            $hyperlinks=$DocumentToEdit.Hyperlinks
            foreach($hyperlink in $hyperlinks) {
                $Address=$hyperlink.Address
                $Address= $Address.Replace($StringToBeReplaced,$NewString) ###this works
                $hyperlink.Address=$Address
            } #end of foreach($hyperlink in $hyperlinks)
        "Saving changes to {0}" -f $DocumentToEdit.Fullname
            $DocumentToEdit.Save()    
            "Completed processing {0} `r`n" -f $DocumentToEdit.Fullname
            $DocumentToEdit.Close()
        }#end of foreach ($Document in $Documents)
     } #end of if($Replace)
     else{
        foreach ($Document in $Documents){
            
            
            
            $DocumentToEdit = $Word.documents.open($Document.FullName,$null, $false, $null, $passwd, $passwdTemplate, $false, $passwdWrite)
            "Processing file: {0}" -f $DocumentToEdit.FullName    
            $hyperlinks=$DocumentToEdit.Hyperlinks
            [string]$InitialDocumentPath=$null
            $InitialDocumentPath=$Document.PSParentPath ###ekei pou einai ta peiramata
            $InitialDocumentPath=$InitialDocumentPath.Replace("Microsoft.PowerShell.Core\FileSystem::","")
            Push-Location $InitialDocumentPath
            foreach($hyperlink in $hyperlinks) {
                if(($hyperlink.Address -like "*http*") -or ($hyperlink.Address -like "*www*")){
                    continue
                }

                $PathInfo=[System.Uri]$hyperlink.Address;
                if($PathInfo.IsUnc){
                #if($hyperlink -like "*\\*"){
                    $SourceIsUNC=$true
                }
                else{
                    $SourceIsUNC=$false
                }


                $Address=$hyperlink.Address
                   $IsFolder=$false
                    $IsFile=$true
                   #if(Test-Path $Address -PathType Leaf){
                       
                      
                         ######to kaname gt to push location den tha paei sosta
                        if(Test-Path $hyperlink.Address -PathType Leaf){
                            $IsFolder=$false
                           # $FileName=(Get-Item -Path $hyperlink.Address).BaseName
                            #$extension=(Get-Item -Path $hyperlink.Address).Extension
                            $FileName=[System.IO.Path]::GetFileNameWithoutExtension("$($hyperlink.Address)")
                            $extension=[System.IO.Path]::GetExtension("$($hyperlink.Address)")
                            $file=[System.IO.Path]::GetFileName("$($hyperlink.Address)")
                            $Address="$Address\.."                        
                            
                        }#if(Test-Path $hyperlink.Address -PathType Leaf)
                        elseif(Test-Path $hyperlink.Address -PathType Container){
                            $IsFolder=$true
                            $IsFile=$false
                        <#$FoldersArray=$Address.Split('\').Split('/')
                        $lastFile=$FoldersArray[$FoldersArray.Length-1]
                        if($lastFile -like "*.*"){
                            $Address="$Address\.."
                        }#>

                        }
                   
                
                [int]$pops=0
                if(Test-Path $Address){
                    Push-Location $Address
                    $pops++
                }
                elseif(Test-Path $InitialDocumentPath){
                
                ####going to the initial source path of document
                   
                    "Change the initial path"
                   $pops++
                    Push-Location $InitialDocumentPath #initial source
                   
                    if(Test-Path $Address){
                        $pops++
                        #3rd pop
                        Push-Location $Address
                    }                   
                    else{ #else in initial source
                   
                        

                            $Address=$Address.Replace("..\","").Replace("../","")
                            if($IsFile){
                                $Address="$Address\.."
                            }

                            $Address="$InitialDocumentPath\$Address"

                            if(Test-Path $Address){
                                $pops++
                                Push-Location $Address

                            }
                            else{
                                <#if($hyperlink.Address -like "*user redirection*"){
                                    $Address=$Address.Replace("$Root","").Replace("\\domainname\user redirection folder2`$\","").Replace("\\domainname\user redirection folder2`$\","")
                                    $Address="$Redirection$Address"
                                }#>
                                $Error.Clear()
                                $pops++
                                Push-Location $Address | Out-Null
                                if($Error){
                                    $pops--
                                    foreach($ErrorItem in $Error){
                                        $ErrorItem.Exception  | Out-File "$destination\errors\$($Document.BaseName).txt" -Append
                                    } 
                            
                                    $Error.Clear()
                                    for($pop=0;$pop -lt $pops;$pop++ ){
                                        Pop-Location
                                    }
                                    "all wrong"
                                    continue #bypass the hyperlink
                                }#if($Error)
                            }
                        
                    }#else in initial source                
                  
                  } #end elseif(Test-Path $InitialDocumentAddress){  #>
                else{
                    $Error.Clear()
                    Push-Location $InitialDocumentPath | Out-Null
                    $pops++
                    if($Error){
                        $pops--
                        foreach($ErrorItem in $Error){
                            $ErrorItem  | Out-File "$destination\errors\$($Document.BaseName).txt" -Append
                        }
                        $Error.Clear()
                        for($pop=0;$pop -lt $pops;$pop++ ){
                            Pop-Location
                        }
                        "all wrong"
                        continue #bypass the hyperlink
                    }

                }   
                  Push-Location -StackName stack1

                  [string]$source=$null
                  $source=Get-Location -StackName stack1
                  $source=$source.replace("Microsoft.PowerShell.Core\FileSystem::","")
                  if(-not $IsFolder){
                    $source="$source\$file"
                  }
                  "Source before copy function :$source"
                  Pop-Location -StackName stack1
                  for($pop=0;$pop -lt $pops;$pop++ ){
                     Pop-Location
                  }
                 # Pop-Location
                 # if($SecondPop){
                  #  Pop-Location
                  #}


                  if(test-path $source -PathType Container){
                    $IsFolder=$true
                    $folder=Split-Path -Path $source -Leaf -Resolve 
                  }
                  elseif(Test-Path $source -PathType Leaf){
                    #$IsFolder=$false
                    $LastDirectory=Split-Path -Path $source -Leaf -Resolve 
                    #$file=(Get-Item -Path $destination3).BaseName
                    #$extension=(Get-Item -Path $destination3).Extension
                    

                  }
                  
                  if($IsFolder){
                    [string]$destination2=$destination+'\'+$Document.BaseName+'\'+$folder
                    $i=0
                    while(Test-Path $destination2){
                        $destination2=$destination+'\'+$Document.BaseName+'\'+$folder+'_'+$i
                        $i++
                    }
                    $i=0
                    }#if($IsFolder){
                  else{
                    [string]$destination2=$destination+'\'+$Document.BaseName+'\'+$FileName+$extension
                    $i=0
                    while(Test-Path $destination2){
                        $destination2=$destination+'\'+$Document.BaseName+'\'+$FileName+'_'+$i+$extension
                        $i++
                    }
                    $i=0
                  }



                    if($Error){
                        foreach($ErrorItem in $Error){
                                    $ErrorItem  | Out-File "$destination\errors\$($Document.BaseName).txt" -Append
                                }
                        $Error.Clear() 
                            continue
                    } #if($Error){
                    
                    
                CopyVeryLongPath -Source $source -Root $DocumentToEdit.FullName -Destination $destination2 -IsFolder:$IsFolder -PathToCreate $destination -DocumentName $($Document.BaseName) -SourceIsUNC:$SourceIsUNC -destinationIsUNC:$destinationIsUNC
                $hyperlink.Address=$destination2
               
                 
                #$hyperlink.Address=$destination+'\'+$LastFileOrFolderOfPath+$global:i
                #$Global:i++ #>
            } #end of foreach($hyperlink in $hyperlinks)

            
             $Report="$Destination\$($DocumentToEdit.Name)"
            "Saving changes to {0}" -f $Report
            $DocumentToEdit.SaveAs2([ref]$Report,[ref]$SaveFormat::wdFormatDocument,$false,$passwd,$false,$passwdWrite) 
            "Completed processing {0} `r`n" -f $Report
            $DocumentToEdit.Close()
            
            Pop-Location #pop location for $InitialDocumentPath
        }#end of foreach ($Document in $Documents){

    }#end of else

    $Word.Quit()
    $Word=$null
    
    if($Both){
            Edit-Excel -Replace:$Replace
    }
    $Description=$null
    $Replace=$false
    $Both=$false
    [gc]::collect()
    [gc]::WaitForPendingFinalizers()
}


do{
    Write-Host "What do you want to do?Press a number to choose from the menu"
    Write-Host "1. Edit hyperlinks"
    Write-Host "2. Copy the target of hyperlinks and move them to antoher directory"
    Write-Host "3. Disable scripting execution, after that you will have to enable in order to run the script!!!!"
    Write-Host "4. Exit"
    [int]$loop1=Read-Host 

    if(($loop1 -le 0) -or ($loop1 -ge 5)){
        Write-Host "Hahaha very funny"
    }
    elseif(($loop1 -eq 1) -or ($loop1 -eq 2)){
        $Replace=$false
        $DoSomething="to copy"
        do{
            
           # Write-Host 
           if($loop1 -eq 1){
            $Replace=$true
            $DoSomething="to replace"
           }
           else{
            $Replace=$false
            $DoSomething="to copy"
           }
            Write-Host "You selected $DoSomething the hyperlinks. Select one of the programs which contain the hyperlinks"
            Write-Host "1. Word"
            Write-Host "2. Excel"
            Write-Host "3. Both"
            Write-Host "4. Exit"
            [int]$loop2=Read-Host
            if($loop2 -eq 1){Edit-Word -Replace:$Replace}#word
            elseif($loop2 -eq 2){Edit-Excel -Replace:$Replace}#excel
            elseif($loop2 -eq 3){Edit-Word -Replace:$Replace -Both}#both
            if(($loop2 -le 0) -or ($loop2 -ge 5)){
                Write-Host "Applause  So funny I forgot to laugh"
            }
            $Replace=$false
        }while($loop2 -ne 4)
    }
    elseif($loop1 -eq 3){
        Set-ExecutionPolicy undefined
    }
}while($loop1 -ne 4)


