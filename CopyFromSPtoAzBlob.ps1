
 Clear-Host
 $OutputEncoding = [Console]::OutputEncoding = New-Object System.Text.Utf8Encoding
 $SPHost = "https://sprcomunicacao.sharepoint.com/"
 $SPSiteUrl = "/sites/SPRCLIENTES/"    
 $SPDocLibraryTitle = "Documentos"
 $azStorageAccountKey = "+yTIIy32SZdOKVODQKzAXUGvrbEWwTQM9qIrGk/MQlWEeTUaAcI7/Wtyok7Wlnm9eZ9xFs7vEGHn+ASt3xqdvA=="
 $azStorageAccountName = "stggrprdsprhistorico"
 $BaseFolderName = "Documentos Partilhados/Clientes2023/"
 $localFileDownloadFolderPath = 'C:\'
 $SPFullSiteUrl = $SPHost + $SPSiteUrl
 $ListFile = ""
 $ActualYear = [int](Get-Date -UFormat %Y)
 $OneYearAgo = [int]((Get-Date).AddYears(-1) | Get-Date -UFormat %Y)
 $OneYearAfter = [int]((Get-Date).AddYears(+1) | Get-Date -UFormat %Y)
 cd  $localFileDownloadFolderPath 
 Clear-Host
 $ErrorActionPreference = 'Stop' 
 $TestYear = {
    try
    {
        write-host 
        write-host 
        write-host "Please, type a year base to copy" -ForegroundColor Blue
        write-host 
        write-host 
        $azStorageContainerName  = [int](Read-Host "Year Base")
        write-host 
        write-host 

        if (($azStorageContainerName -lt $OneYearAgo) -or ($azStorageContainerName -gt $OneYearAfter)) {
            Write-Host "Your input has to be a number greater or equal than $OneYearAgo !" -ForegroundColor Red
            write-host 
            write-host 
            & $TestYear
        }
        else {
            $azStorageContainerName
        }
    }
    catch
    {
        write-host 
        write-host 
        Write-Host "Your input has to valid range year." -ForegroundColor Red
        write-host 
        write-host 
        & $TestYear
    }
}

$userInput = & $TestYear
$azStorageContainerName =  $userInput

start-sleep 3
write-host $azStorageContainerName 
write-host "Please, Specify if is a Full Initial Copy or is a Partial Copy (if partial you need to fill date from start in next steps." -ForegroundColor Blue
$CopyType = Read-host "Input your CopyType: (F to Full Copy or P to Partial Copy)" 


if ( ($CopyType -ne "P") -and ($CopyType -ne "F") ) {
 
    write-host "Wrong Parameter" -ForegroundColor Red
    exit
}

start-sleep 3

if ($CopyType -eq 'F') { 

    $BaseFolderName = "Documentos Partilhados/Clientes" + $azStorageContainerName + "/"
    $SPfolderFirstLevel = m365 spo folder list --webUrl $SPFullSiteUrl --parentFolderUrl  $BaseFolderName --fields 'ServerRelativeUrl,Name' -o json | ConvertFrom-json
    
    ForEach ($i in $SPfolderFirstLevel) {
    $FolderUrl = $SPSiteUrl + $BaseFolderName + $i.Name
    $SPLibraryItems = m365 spo file list --webUrl $SPFullSiteUrl  --folderUrl $FolderUrl  --fields 'ServerRelativeUrl,Name' -r -o json | ConvertFrom-Json 

         if (($SPLibraryItems.Count -gt 0) -and ($SPLibraryItems.ServerRelativeUrl -cnotcontains '@Agencia/SPR_Corporativo')) {

             ForEach ($SPLibraryItem in $SPLibraryItems) {
                $SPLibFileRelativeUrl = $SPLibraryItem.ServerRelativeUrl
                $SPFileName = $SPLibraryItem.ServerRelativeUrl
                $SPLibraryFolderRelativeUrl = $SPLibFileRelativeUrl.Substring(0, $SPLibFileRelativeUrl.lastIndexOf('/'))
                $SPFileName = $SPLibraryItem.ServerRelativeUrl
                $SPParentDir = $SPLibraryItem.ServerRelativeUrl
                $FolderParent = $FolderUrl + $SPParentDir
                New-Item -ItemType Directory -Force -Path  '$localFileDownloadFolderPath/$azStorageContainerName/$FolderUrl/$FolderParent' | Out-Null
                $FileName = $SPLibraryItem.Name
                $SPLibFileRelativeUrl = $SPLibraryItem.ServerRelativeUrl
                $SPFileName = $SPLibraryItem.ServerRelativeUrl
                $SPLibraryFolderRelativeUrl = $SPLibFileRelativeUrl.Substring(0, $SPLibFileRelativeUrl.lastIndexOf('/'))
                $SPFileName = $SPLibraryItem.ServerRelativeUrl
                $localDownloadFolderPath = $SPLibraryItem.ServerRelativeUrl
                $SPDestFileName = $SPFileName.split("/", 5)[4]
                $SPLibraryFolderRelativeUrl = $SPDestFileName.Substring(0, $SPDestFileName.lastIndexOf('/'))
                $localSourceFilePath = Join-Path $azStorageContainerName $SPLibraryFolderRelativeUrl
                $SPLibFileDestRelativeUrl = Join-path  $azStorageContainerName $SPDestFileName 
 
                    If (!(test-path $localSourceFilePath)) {
                        $message = "Target local cache folder  $localSourceFilePath not exist"
                        Write-Host $message -ForegroundColor Yellow
                    
                    New-Item -ItemType Directory -Force -Path   $localSourceFilePath  | Out-Null
                
                        $message = "Created target local cache folder at $localSourceFilePath"
                        Write-Host $message -ForegroundColor Green
                    }
                        else {
                            $message = "Target local cache folder exist at  $localSourceFilePath"
                            Write-Host $message -ForegroundColor Blue
                    }

                $localFilePath =  Join-path $azStorageContainerName $SPLibraryFolderRelativeUrl
                $message = "Processing SharePoint file $SPFileName"
                Write-Host $message -ForegroundColor Green
          
                Clear-Variable -Name ListFile
                $FileonBlob = $SPLibraryFolderRelativeUrl + '/' + $FileName
                $ListFile = az storage blob list --account-name $azStorageAccountName --account-key  "$azStorageAccountKey" -c $azStorageContainerName --delimiter '/' --query '[].name' --prefix $FileonBlob --only-show-errors -o json | ConvertFrom-Json
 
                if ( [string]::IsNullOrEmpty($ListFile)) {
                    Write-Host New File or  Replaced updated File -ForegroundColor Green
                    m365 spo file get --webUrl $SPFullSiteUrl --url $SPLibFileRelativeUrl --asFile --path  $SPLibFileDestRelativeUrl | Out-Null

                    $message = "Downloaded SharePoint file at $localFilePath" 
                    Write-Host $message -ForegroundColor Green
                    $localFolderToSync = Join-Path $azStorageContainerName $SPLibraryFolderRelativeUrl | Out-Null

                    $message = "Target local folder  $localSourceFilePath not exist"
                    az storage blob sync --account-name $azStorageAccountName --account-key  "$azStorageAccountKey" -c $azStorageContainerName  -s $SPLibFileDestRelativeUrl -d $SPLibraryFolderRelativeUrl/$FileName --delete-destination true --only-show-errors  -o json | Out-Null

                    Remove-Item -Force -Path  $SPLibFileDestRelativeUrl | Out-Null
                
                    $message = "Syncing local folder $localFolderToSync with Azure Storage Container $azStorageContainerName is completed"
                    Write-Host $message -ForegroundColor Green
                    }
                        else {
                            write-host FILE $ListFile ALREADY EXIST, ignoring -ForegroundColor Red
                            Clear-Variable -Name ListFile
                    }
             }

        }


    }
}

else {
   
            write-host "Input your StartDate - Format must be dd/mm/aaaa - Sample: 01/12/2023" -ForegroundColor Blue
            write-host 

        do {
                $date = $null
                $today = Read-Host -Prompt ('Enter start date(e.g. {0}) ' -f (Get-Date -Format "dd/MM/yyyy"))
            
                try {
                    $date = Get-Date -Date $today -Format "dd/MM/yyyy" -ErrorAction Stop
                    '{0} is a valid date' -f $date
                }
                catch {
                    '{0} is an invalid date' -f $today
                }
            }
            until ($date)
            $copyFromDate = $date



                $copyFromDateFull = [datetime]::parseexact($copyFromDate, 'dd/MM/yyyy', $null).ToString('yyyy-MM-ddTHH:mm:ss')

                $BaseFolderName = "Documentos Partilhados/Clientes" + $azStorageContainerName + "/"

                $SPfolderFirstLevel = m365 spo folder list --webUrl $SPFullSiteUrl --parentFolderUrl  $BaseFolderName --fields 'ServerRelativeUrl,Name' -o json | ConvertFrom-json
            ForEach ($i in $SPfolderFirstLevel) {
            $FolderUrl = $SPSiteUrl + $BaseFolderName + $i.Name
            $SPLibraryItems = m365 spo file list --webUrl $SPFullSiteUrl  --folderUrl $FolderUrl  --fields 'ServerRelativeUrl,Name,TimeLastModified' -r  --query "[?TimeLastModified > '$($copyFromDateFull)']" -o json | ConvertFrom-Json 
   
                 if ($SPLibraryItems.Count -gt 0) {

                        ForEach ($SPLibraryItem in $SPLibraryItems) {
                        $SPLibFileRelativeUrl = $SPLibraryItem.ServerRelativeUrl
                        $SPFileName = $SPLibraryItem.ServerRelativeUrl
                        $SPLibraryFolderRelativeUrl = $SPLibFileRelativeUrl.Substring(0, $SPLibFileRelativeUrl.lastIndexOf('/'))
                        $SPFileName = $SPLibraryItem.ServerRelativeUrl
                        $SPParentDir = $SPLibraryItem.ServerRelativeUrl
                        $FolderParent = $FolderUrl + $SPParentDir

                        New-Item -ItemType Directory -Force -Path  '$localFileDownloadFolderPath/$azStorageContainerName/$FolderUrl/$FolderParent' | Out-Null
    
                        $FileName = $SPLibraryItem.Name
                        $SPLibFileRelativeUrl = $SPLibraryItem.ServerRelativeUrl
                        $SPFileName = $SPLibraryItem.ServerRelativeUrl
                        $SPLibraryFolderRelativeUrl = $SPLibFileRelativeUrl.Substring(0, $SPLibFileRelativeUrl.lastIndexOf('/'))
                        $SPFileName = $SPLibraryItem.ServerRelativeUrl

                        $localDownloadFolderPath = $SPLibraryItem.ServerRelativeUrl
                        $SPDestFileName = $SPFileName.split("/", 5)[4]
                        $SPLibraryFolderRelativeUrl = $SPDestFileName.Substring(0, $SPDestFileName.lastIndexOf('/'))
                        $localSourceFilePath = Join-Path $azStorageContainerName $SPLibraryFolderRelativeUrl
                        $SPLibFileDestRelativeUrl = Join-path  $azStorageContainerName $SPDestFileName 
 

                            If (!(test-path $localSourceFilePath)) {
                                $message = "Target local folder  $localSourceFilePath not exist"
                                Write-Host $message -ForegroundColor Yellow
                            
                                New-Item -ItemType Directory -Force -Path   $localSourceFilePath  | Out-Null
                        
                                $message = "Created target local folder at $localSourceFilePath"
                                Write-Host $message -ForegroundColor Green
                            }
                            else {
                                $message = "Target local folder exist at  $localSourceFilePath"
                                Write-Host $message -ForegroundColor Blue
                            }

                            $localFilePath =  Join-path $azStorageContainerName $SPLibraryFolderRelativeUrl
                            $message = "Processing SharePoint file $SPFileName"
                            Write-Host $message -ForegroundColor Green
                            
                            Clear-Variable -Name ListFile

                            $FileonBlob = $SPLibraryFolderRelativeUrl + '/' + $FileName
                            
                            Clear-Variable -Name ListFile
                        
                            Write-Host New File or  Replaced File -ForegroundColor Green
                            m365 spo file get --webUrl $SPFullSiteUrl --url $SPLibFileRelativeUrl --asFile --path  $SPLibFileDestRelativeUrl | Out-Null

                            $message = "Downloaded SharePoint file at $localFilePath" 
                            Write-Host $message -ForegroundColor Green
                            $localFolderToSync = Join-Path $azStorageContainerName $SPLibraryFolderRelativeUrl | Out-Null

                            $message = "Target local folder  $localSourceFilePath not exist"
                            az storage blob sync --account-name $azStorageAccountName --account-key  "$azStorageAccountKey" -c $azStorageContainerName  -s $SPLibFileDestRelativeUrl -d $SPLibraryFolderRelativeUrl/$FileName --delete-destination true --only-show-errors  -o json | Out-Null

                            Remove-Item -Force -Path  $SPLibFileDestRelativeUrl | Out-Null
    
                            $message = "Syncing local folder $localFolderToSync with Azure Storage Container $azStorageContainerName is completed"
                            Write-Host $message -ForegroundColor Green


                        }

                }
            }


        }
    
    