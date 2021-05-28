Clear-Host


Function MiniMe-GetEmailReadStatus
{
    param (
            $sender,
            $subject,
            $mailbox,
            $userName,
            $password
          )
	Process{
                $SQ = "From:`"$Sender`" AND Subject:`"$subject`""
                $report=@()
                $itemsView=1000
                $uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
                $dllpath = "C:\Microsoft.Exchange.WebServices.dll"
                Import-Module $dllpath

                $pass=$password
                $AccountWithImpersonationRights=$userName
                $MailboxToImpersonate=$mailbox

                ## Set Exchange Version
                $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
                $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
                $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $AccountWithImpersonationRights, $pass
                $service.url = $uri

                Write-Host 'Using ' $AccountWithImpersonationRights ' Account to work in ' $MailboxToImpersonate
                $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
                ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$MailboxToImpersonate);

                $Folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$ImpersonatedMailboxName)
            
                #$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Folderid)

                $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
                $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $view.PropertySet = $propertyset

                $items = $service.FindItems($Folderid,$SQ,$view)

                if ($items -ne $null){
                    foreach($item in $items){
                        $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                        $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $item.ID, $emailProps)   
                        $datam=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={"Inbox"}}
                        $report+=$datam
                    }
                }                                                                                                                                                                                                                                                                                                                                         Else{
                    Write-Host "Mail Not Found in Inbox Folder for:"$mailbox -f Yellow -NoNewline
                    Write-Host " Checking Deleted Item Folder"
                    $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
                    ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$ImpersonatedMailboxName)
                    $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                    $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                    $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                    $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                    $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note"} | select ID,Displayname
                    $DI = $allFolders | ? {$_.DisplayName -eq "Deleted Items"}
                    $Folderid=$DI.ID
                    $items = $service.FindItems($Folderid,$SQ,$view)
    
                     if ($items.count -eq $null){
                        write-host "Item not found in the Deleted item folder, Now Checking in the Recover Deleted items Folder"
                        $itemsView=90000
                        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
                        $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                        $view.PropertySet = $propertyset    
                        $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
                        ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Recoverableitemsroot,$ImpersonatedMailboxName)
                        $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                        $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                        $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                        $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                        $Deletions = $findFolderResults | ? {$_.DisplayName -eq "Deletions"}
                        $Folderid=$Deletions.ID
                        $items=$service.FindItems($Folderid,$SQ,$view)

                        if ($items.count -eq $null){
                            Write-Host "Item Not Found in the Dumpsters."
                            Write-host "Checking in other folders"
                
                            $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
                            ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$ImpersonatedMailboxName)
                            $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                            $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                            $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                            $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                            $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note"} | select ID,Displayname
                            $allFolders=$allFolders | ? { `
                            $_.DisplayName -notlike "Inbox" -and `
                            $_.DisplayName -notlike "Deleted Items" -and `
                            $_.DisplayName -notlike "Drafts" -and `
                            $_.DisplayName -notlike "Sent Items" -and `
                            $_.DisplayName -notlike "Outbox"}    
                            $allfoldersCount=$allfolders.count
                            $counter=0
                            $itemFound=$false

                            if ($allFolders){
	                            do{
                                        Write-Host "Checking Email Item in Folder:"$allfolders[$counter].DisplayName     
                                        $folderID=$allfolders[$counter].ID     
                                        $items =$service.FindItems($Folderid,$SQ,$view)         
                                        if ($items.count -eq $null){
                                            Write-Host "Item Was Not Found in Folder:"$allfolders[$counter].DisplayName -ForegroundColor Yellow
                                        }   
                                        else{
                                         #   foreach($item in $items){
                                                Write-Host "Item Was Found in Folder:"$allfolders[$counter].DisplayName -ForegroundColor Green     
                                                $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)     
                                                $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $item.ID, $emailProps)      
                                                $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={$allfolders[$counter].DisplayName}}     
                                                $report+=$data         
                                                $itemFound=$true
                                         #   }
                                        } 
                                        $counter++ 
                                    } 
                                    until ($counter -eq $allfoldersCount -or $itemFound -eq $true) 
                            }
                        }             
                        else{
                                $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)             
                                $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps)
                                $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={"Deletions"}}
                                $report+=$data
                        }
                    }          
                    else{        
                        $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
                        $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps)        
                        $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={$DI.DisplayName}}       
                        $report+=$data
                    }
             } 
                $report  
            }
}

Function MiniMe-GetEmailReadStatus
{
    param (
            $sender,
            $subject,
            $mailbox,
            $userName,
            $date,
            $password
          )
	Process{
                $SQ = "From:`"$Sender`" AND Subject:`"$subject`""
                $report=@()
                $itemsView=1000
                $uri=[system.URI] "https://outlook.office365.com/ews/exchange.asmx"
                $dllpath = "C:\Microsoft.Exchange.WebServices.dll"
                Import-Module $dllpath

                $pass=$password
                $AccountWithImpersonationRights=$userName
                $MailboxToImpersonate=$mailbox

                ## Set Exchange Version
                $ExchangeVersion = [Microsoft.Exchange.WebServices.Data.ExchangeVersion]::Exchange2016
                $service = New-Object Microsoft.Exchange.WebServices.Data.ExchangeService($ExchangeVersion)
                $service.Credentials = New-Object Microsoft.Exchange.WebServices.Data.WebCredentials -ArgumentList $AccountWithImpersonationRights, $pass
                $service.url = $uri

                Write-Host 'Using ' $AccountWithImpersonationRights ' Account to work in ' $MailboxToImpersonate
                $service.ImpersonatedUserId = New-Object Microsoft.Exchange.WebServices.Data.ImpersonatedUserId `
                ([Microsoft.Exchange.WebServices.Data.ConnectingIdType]::SMTPAddress,$MailboxToImpersonate);

                $Folderid= new-object Microsoft.Exchange.WebServices.Data.FolderId([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Inbox,$ImpersonatedMailboxName)
            
                #$MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$Folderid)

                $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
                $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                $view.PropertySet = $propertyset

                $items = $service.FindItems($Folderid,$SQ,$view)

                if ($items -ne $null){
                    foreach($item in $items){
                        $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)
                        $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $item.ID, $emailProps)   
                        $datam=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={"Inbox"}}
                        $report+=$datam
                    }
                }                                                                                                                                                                                                                                                                                                                                         Else{
                    Write-Host "Mail Not Found in Inbox Folder for:"$mailbox -f Yellow -NoNewline
                    Write-Host " Checking Deleted Item Folder"
                    $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
                    ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$ImpersonatedMailboxName)
                    $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                    $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                    $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                    $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                    $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note"} | select ID,Displayname
                    $DI = $allFolders | ? {$_.DisplayName -eq "Deleted Items"}
                    $Folderid=$DI.ID
                    $items = $service.FindItems($Folderid,$SQ,$view)
    
                     if ($items.count -eq $null){
                        write-host "Item not found in the Deleted item folder, Now Checking in the Recover Deleted items Folder"
                        $itemsView=90000
                        $view = New-Object Microsoft.Exchange.WebServices.Data.ItemView -ArgumentList $ItemsView
                        $propertyset = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::IdOnly)
                        $view.PropertySet = $propertyset    
                        $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
                        ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::Recoverableitemsroot,$ImpersonatedMailboxName)
                        $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                        $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                        $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                        $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                        $Deletions = $findFolderResults | ? {$_.DisplayName -eq "Deletions"}
                        $Folderid=$Deletions.ID
                        $items=$service.FindItems($Folderid,$SQ,$view)

                        if ($items.count -eq $null){
                            Write-Host "Item Not Found in the Dumpsters."
                            Write-host "Checking in other folders"
                
                            $MailboxRootid= new-object Microsoft.Exchange.WebServices.Data.FolderId `
                            ([Microsoft.Exchange.WebServices.Data.WellKnownFolderName]::MsgFolderRoot,$ImpersonatedMailboxName)
                            $MailboxRoot=[Microsoft.Exchange.WebServices.Data.Folder]::Bind($service,$MailboxRootid)
                            $FolderList = new-object Microsoft.Exchange.WebServices.Data.FolderView(100)
                            $FolderList.Traversal = [Microsoft.Exchange.WebServices.Data.FolderTraversal]::Deep
                            $findFolderResults = $MailboxRoot.FindFolders($FolderList)
                            $allFolders=$findFolderResults | ? {$_.FolderClass -eq "IPF.Note"} | select ID,Displayname
                            $allFolders=$allFolders | ? { `
                            $_.DisplayName -notlike "Inbox" -and `
                            $_.DisplayName -notlike "Deleted Items" -and `
                            $_.DisplayName -notlike "Drafts" -and `
                            $_.DisplayName -notlike "Sent Items" -and `
                            $_.DisplayName -notlike "Outbox"}    
                            $allfoldersCount=$allfolders.count
                            $counter=0
                            $itemFound=$false

                            if ($allFolders){
	                            do{
                                        Write-Host "Checking Email Item in Folder:"$allfolders[$counter].DisplayName     
                                        $folderID=$allfolders[$counter].ID     
                                        $items =$service.FindItems($Folderid,$SQ,$view)         
                                        if ($items.count -eq $null){
                                            Write-Host "Item Was Not Found in Folder:"$allfolders[$counter].DisplayName -ForegroundColor Yellow
                                        }   
                                        else{
                                         #   foreach($item in $items){
                                                Write-Host "Item Was Found in Folder:"$allfolders[$counter].DisplayName -ForegroundColor Green     
                                                $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)     
                                                $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $item.ID, $emailProps)      
                                                $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={$allfolders[$counter].DisplayName}}     
                                                $report+=$data         
                                                $itemFound=$true
                                         #   }
                                        } 
                                        $counter++ 
                                    } 
                                    until ($counter -eq $allfoldersCount -or $itemFound -eq $true) 
                            }
                        }             
                        else{
                                $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties)             
                                $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps)
                                $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={"Deletions"}}
                                $report+=$data
                        }
                    }          
                    else{        
                        $emailProps = New-Object Microsoft.Exchange.WebServices.Data.PropertySet ([Microsoft.Exchange.WebServices.Data.BasePropertySet]::FirstClassProperties) 
                        $mail = [Microsoft.Exchange.WebServices.Data.EmailMessage]::Bind($service, $items.ID, $emailProps)        
                        $data=$mail | Select @{N="Sender";E={$_.Sender.Address}},@{N="USER";E={$mailbox}},Isread,Subject,DateTimeReceived,@{n="Folder";e={$DI.DisplayName}}       
                        $report+=$data
                    }
             } 
                $report  
            }
}

Function MiniMe-O365-Connect{
	Param (
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][PsObject]$Credentialobj,
		)
	Process {
		Import-Module "MSOnline" -Global
		$Session= New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/powershell-liveid/' -Credential $Credentialobj -Authentication Basic –AllowRedirection
		if ($Session -ne $null){
			Import-Module (Import-PSSession -Session $Session -AllowClobber -CommandName * -DisableNameChecking) -Global
			Import-PSSession -Session $Session -AllowClobber -CommandName * -DisableNameChecking 
			Connect-MsolService -Credential $Credentialobj
			return $true
		}else{
			return $false
		}
		
	}
}

Function MiniMe-CredentialObj{
	Param (
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$UserName,
		[Parameter(Mandatory=$true)][ValidateNotNullorEmpty()][string]$Password,
		)
	Process {
		$Secpwd = ConvertTo-SecureString $Password -AsPlainText -Force
		$Cred = New-Object System.Management.Automation.PSCredential ($UserName, $secpwd)
		Return , $Cred
	}
}

Function MiniMe-Logme{
	param(
		[String]$Text,
		[PsObject]$ErrorObj,
		[ValidateSet('Black', 'DarkBlue', 'DarkGreen', 'DarkCyan', 'DarkRed', 'DarkMagenta', 'DarkYellow', 'Gray', 'DarkGray', 'Blue', 'Green', 'Cyan', 'Red', 'Magenta', 'Yellow', 'White')][String] $TextColor,
		[PsObject]$ExecuteTime,
		[String]$LogFileName,
		[String]$ToEmail,
		[String]$FromEmail,
		[String]$SMTPserver
		)
	Process {
		if ($ErrorObj -ne $null){
			$Text = 'ERROR: ' + $ErrorObj[0] + '-------' + "`n" + $Text
			$ErrorObj.clear()
		}

		if($Text -ne '' -or $ErrorObj -ne $null){
			$ExeTime=''
			if ($ExecuteTime -ne $null){$ExeTime = ' - EXEtime: ' + "$($ExecuteTime.Days):$($ExecuteTime.Hours):$($ExecuteTime.Minutes):$($ExecuteTime.Seconds):$($ExecuteTime.Milliseconds)"}
			
			$FullLog  =(Get-Date -Format "MM/dd/yyyy HH:mm:ss") + " " + $Text + $ExeTime
			
			if ($LogFileName -ne ''){$FullLog | Out-File $LogFileName -Append}

			if ($TextColor -ne ''){
				write-host $FullLog -ForegroundColor $TextColor
			}else{
				write-host $FullLog
			}
			if ($ToEmail -ne ''){
				#if($FromEmail -eq ''){$FromEmail = 'Kraken Reporting <Kraken@Consulatehc.com>'}
				RW-SendMail -FromEmail $FromEmail -ToEmail $ToEmail -SMTPserver $SMTPserver -Subject 'Log Notificaton' -Msg $FullLog -LogFilePath $LogFileName
			}
		}
	}
}

$o365_Global_Admin_User = '' # Use your o365 tenant global admin account  
$o365_Global_Admin_Pwd = '' # Use your o365 tenant global admin account  

$O365CredObj = MiniMe-CredentialObj -UserName $o365_Global_Admin_User -Password $o365_Global_Admin_Pwd
MiniMe-Logme -Text "Creating o365 Cred obj" -LogFileName $ScriptLog

$Subject = '' # Campaign Email Subject
$Sender = '' # Campaign Sender email address

$command=$null; $command = [scriptblock]::Create('Get-MessageTrace -SenderAddress ' + $Sender  + ' -RecipientAddress ' + $Sender)
$MsgID=$null; $MsgID = Invoke-Command -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/powershell' -Credential $O365CredObj -Authentication Basic -AllowRedirection -ScriptBlock $command | where {$_.subject -eq $subject}
MiniMe-Logme -LogFileName $ScriptLog -Text "--- Results: $($result)" 


if($MsgID){
    $MsgTrace=$null; $Error.Clear()

    For($i = 1; $i -le 1000; $i++){
		$command=$null; $command = [scriptblock]::Create('Get-MessageTrace -MessageId ' + (($MsgID.MessageId).replace('<','')).replace('>','') +' -PageSize 5000 -Page ' + $i)
		$MsgTraceRslt=$null;$MsgTraceRslt = Invoke-Command -ConfigurationName Microsoft.Exchange -ConnectionUri 'https://outlook.office365.com/powershell' -Credential $O365CredObj -Authentication Basic -AllowRedirection -ScriptBlock $command
		if($MsgTraceRslt.count){
           $MsgTrace += $MsgTraceRslt
        }
        else{
            break
        }		
    }

    MiniMe-Logme -Text "Msg Count: $($MsgTrace.count)" -LogFileName $ScriptLog

    $MsgTraceDump = @(); $count = 0
    
    foreach ($Email in $MsgTrace){
        Write-Progress -Activity "Email: $($Email.RecipientAddress)" -PercentComplete (($count/$MsgTrace.count)*100)
        $UPN =$Email.RecipientAddress

        $ADUserInfo=$null; $ADUserInfo = Get-ADUser -Filter {mail -eq $UPN}
        $Rsult = MiniMe-GetEmailReadStatus -sender $Sender -subject $subject -mailbox $o365_Global_Admin_User -password $o365_Global_Admin_Pwd

        if($ADUserInfo){
            $rpt = New-Object -TypeName psobject
            $rpt | Add-Member -MemberType NoteProperty -Name Email -Value $Email.RecipientAddress

            if(($Rsult.IsRead -ne $null) -and ($Rsult.IsRead -ne '')){
                $rpt | Add-Member -MemberType NoteProperty -Name IsRead -Value $Rsult.IsRead
            }else{
                $rpt | Add-Member -MemberType NoteProperty -Name IsRead -Value 'Unknown'
            }
            $rpt | Add-Member -MemberType NoteProperty -Name Folder -Value $Rsult.Folder
            $rpt | Add-Member -MemberType NoteProperty -Name Status -Value $Email.Status
        }else{
            $rpt = New-Object -TypeName psobject
            $rpt | Add-Member -MemberType NoteProperty -Name Email -Value $Email.RecipientAddress
            if(($Rsult.IsRead -ne $null) -and ($Rsult.IsRead -ne '')){
                $rpt | Add-Member -MemberType NoteProperty -Name IsRead -Value $Rsult.IsRead
            }else{
                $rpt | Add-Member -MemberType NoteProperty -Name IsRead -Value 'Unknown'
            }
            $rpt | Add-Member -MemberType NoteProperty -Name Folder -Value $Rsult.Folder
            $rpt | Add-Member -MemberType NoteProperty -Name Status -Value $Email.Status
            $rpt | Add-Member -MemberType NoteProperty -Name JobCode -Value 'N/A'
        }
        $MsgTraceDump += $rpt
        $count++
    }
    
    $SMTP =''
    $Sender='' #email sender address 
    $Reciever='' 
    $Subject = 'eMail Analytics'
    
    $MainInfo = $MsgTrace | Group-Object Subject | select Name,count
    $IsReadRaw=$null;$IsReadRaw = $MsgTraceDump | Group-Object IsRead | select @{Name = "Is Read"; Expression = {$_.Name}}, @{Name = "Total Count"; Expression = {$_.Count}}
    $IsDeliveredRaw=$null;$IsDeliveredRaw = $MsgTraceDump | where{($_.Status -eq 'Delivered') -or $_.Status -eq 'Failed'} | Group-Object Status | select @{Name = "Status"; Expression = {$_.Name}}, @{Name = "Total Count"; Expression = {$_.Count}}

    $HTMLPreContent=$null; $HTMLReport=$null;$HTMLHead=$null
    $HTMLPreContent	= "<Center><h3><font color=#294963>eMail Analytics $(get-date -Format "MM/dd/yyyy")</font></h3><p><b> Campaign: $($MainInfo.Name)</b></p>"
    $HTMLHead = "<style>$(RW-CssStyle -Style table -LogFilePath $ScriptLog)</style>"
    $HTMLReport +=  ConvertTo-Html -Head $HTMLHead -PreContent $HTMLPreContent | Out-String
    
    $HTMLReport += "<P><Center><P><b>Read status</b>"
    $HTMLReport +=  $IsReadRaw | select * | ConvertTo-Html | Out-String

    $HTMLReport += "<P><Center><P><b>Delivery status</b>"
    $HTMLReport +=  $IsDeliveredRaw | select * | ConvertTo-Html | Out-String

    $HTMLReport += "<P><Center><P><b>Status Report</b>"
    $HTMLReport += $MsgTraceDump | select * | ConvertTo-Html | Out-String

    Send-MailMessage -From $Sender -BodyAsHtml -Body $HTML -to $Reciever -Subject $Subject -Msg $HTMLReport -SmtpServer $SMTP

    $Report | export-csv "$(split-path -parent $MyInvocation.MyCommand.Definition)\$(get-date -Format MM-dd-yyyy).csv" -NoTypeInformation
}
