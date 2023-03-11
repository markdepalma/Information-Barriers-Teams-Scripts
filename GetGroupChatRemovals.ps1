Param(
	[Parameter(Mandatory=$true)]
	[string] $CsvPath,
	
	[Parameter(Mandatory=$true)]
	[string] $TenantId,
	
	[Parameter(Mandatory=$true)]
	[string] $AppId,
	
	[Parameter(Mandatory=$true)]
	[string] $AppSecret,
	
	[Parameter(Mandatory=$true)]
	[string] $StartDateTimeString,
	
	[Parameter(Mandatory=$true)]
	[string] $EndDateTimeString,
	
	[Parameter(Mandatory=$true)]
	[string] $OutputCsvPath,
	
	[Parameter(Mandatory=$true)]
	[string] $ErrorTxtPath,
	
	[Parameter(Mandatory=$false)]
	[switch] $IncludeMeetings
)

Import-Module MSAL.PS

$StartDateTime = [datetime]$StartDateTimeString
$EndDateTime = [datetime]$EndDateTimeString

#HTTP Retry Settings
$HttpRetries = 3
$HttpRetrySleepSeconds = 10

$ChatFilterString = "chatType eq 'group'"

If ($IncludeMeetings -eq $true) {
	$ChatFilterString = $ChatFilterString + " or chatType eq 'meeting'"
}

$RemovalList = @()
$ErrorUsers = @()

$UserList = @(Import-Csv -Path $CsvPath)




Function Connect-TeamsAPI {
	If ($Token.ExpiresOn.LocalDateTime -gt (Get-Date).AddMinutes(15)) {
		Return
	}
	Else {
		Write-Host '-Refreshing Teams API access token...'
	}
	
	$AppSecretString = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
	$global:Token = Get-MsalToken -ClientId $AppId -ClientSecret $AppSecretString -TenantId $TenantId -ForceRefresh
	$global:Headers = @{Authorization = "Bearer $($Token.AccessToken)"}
}

Function Invoke-GraphRequest {
	Param(
		[Parameter(Mandatory=$true)]
		[string] $Uri,
		
		[Parameter(Mandatory=$true)]
		[object] $Headers,
		
		[Parameter(Mandatory=$false)]
		[string] $Method = 'GET',
		
		[Parameter(Mandatory=$false)]
		[string] $Body = ''
	)
	
	$Response = $null
	$Tries = 0
	
	$Params = @{Uri = $Uri; Headers = $Headers; Method = $Method}
	
	If ($Body -ne '') {
		$Params = $Params + @{Body = $Body}
	}
	
	Do {
		$Tries = $Tries + 1
		
		Try {
			$Response = Invoke-WebRequest @Params -ErrorAction Stop #-Verbose
			$ResponseJson = ($Response | ConvertFrom-Json)
			Return $ResponseJson
		}
		Catch {
			Write-Host "-HTTP Error: $($_.Exception.Response.StatusCode.value__)"
			
			If ($_.Exception.Response.StatusCode.value__ -eq 401 -or $_.Exception.Response.StatusCode.value__ -eq 404) {
				Return
			}
			
			If ($Tries -le $HttpRetries) {
				Write-Host "--Sleeping $HttpRetrySleepSeconds seconds before retrying..."
				Start-Sleep -Seconds $HttpRetrySleepSeconds
			}
		}
	} While ($Tries -le $HttpRetries)
	
	Write-Host "--Max retries reached!"
	
	Return $ResponseJson
}




#Progress counters
#$a = Users
#$b = GetChatIterations
#$c = IterateChats
#$d = GetMessages
#$f = UserSpecificRemovals

$a = 1

ForEach ($User in $UserList) {
	$f = 0
	
	Write-Progress -Id 0 -Activity "Processing user ($a / $($UserList.Count)): $($User.UserPrincipalName)" -Status "Removals found under user: $f  |  Total Removals: $($RemovalList.Count)" -PercentComplete (($a / $UserList.Count) * 100)
	
	Connect-TeamsAPI
	
	Write-Host "Processing user: $($User.UserPrincipalName)..."
	
	$Chats = @()
	$ChatResults = @()
	
	$Url = "https://graph.microsoft.com/beta/users/$($User.UserPrincipalName)/chats?`$filter=$ChatFilterString&`$expand=members"
	
	$b = 1
	
	Do {
		Connect-TeamsAPI
		
		#Write-Host "Getting: $Url"
		
		Write-Progress -Id 1 -ParentId 0 -Activity "Getting chats" -Status "Page: $b | Chat count: $($Chats.Count)"
		
		$ResponseJson = $null
		$ResponseJson = Invoke-GraphRequest -Uri $Url -Headers $Headers
		
		If ($ResponseJson -ne $null) {
			$ChatResults = @($ResponseJson.Value)
			
			$ExceededStartDateTime = $true
			#$ChatResults.Count
			ForEach ($ChatResult in $ChatResults) {
				If ([datetime]($ChatResult.lastUpdatedDateTime) -ge $StartDateTime) {
					$ExceededStartDateTime = $false
					$Chats = $Chats + $ChatResult
				}
			}
			
			$b = $b + 1
			
			$Url = $ResponseJson.'@odata.nextLink'
		}
		Else {
			"-Error getting chats for user"
			$ErrorUsers = $ErrorUsers + $User.UserPrincipalName
			Break
		}
	} While ($Url -ne $null -and $ExceededStartDateTime -eq $false)
	
	Write-Host "-Evaluating chat count: $($Chats.Count)"
	
	$c = 1
	
	ForEach ($Chat in $Chats) {
		Connect-TeamsAPI
		
		Write-Progress -Id 1 -ParentId 0 -Activity "Reading chats" -Status "Chat $c of $($Chats.Count)" -PercentComplete (($c / $Chats.Count) * 100)
		
		$Url = "https://graph.microsoft.com/beta/users/$($User.UserPrincipalName)/chats/$($Chat.id)/messages?`$orderBy=createdDateTime desc"
		
		$d = 1
		
		$LastChatMessageActivity = [datetime]'0001-01-01T00:00:00Z'
		$TempRemovalList = @()
		
		Do {
			Write-Progress -Id 2 -ParentId 1 -Activity "Getting messages" -Status "Page: $d"
			
			$ResponseJson = $null
			$MessageResults = @()
			
			$ResponseJson = Invoke-GraphRequest -Uri $Url -Headers $Headers
			$MessageResults = @($ResponseJson.Value)
			
			If ($ResponseJson -eq $null) {
				Write-Host "-Error getting chat id: $($Chat.id) | chat topic: $($Chat.topic)... Possible external chat"
				$Url = $null
			}
			
			Write-Progress -Id 3 -ParentId 2 -Activity "Reading $($MessageResults.Count) messages"
			
			ForEach ($MessageResult in $MessageResults) {
				Connect-TeamsAPI
				
				If ($MessageResult.messageType -eq 'message' -and $MessageResult.lastModifiedDateTime -gt $LastChatMessageActivity) {
					$LastChatMessageActivity = [datetime]($MessageResult.lastModifiedDateTime)
				}
				
				If ($MessageResult.messageType -eq 'systemEventMessage' -and $MessageResult.eventDetail.'@odata.type' -eq '#microsoft.graph.membersDeletedEventMessageDetail' -and [datetime]$MessageResult.createdDateTime -ge $StartDateTime -and [datetime]$MessageResult.createdDateTime -lt $EndDateTime) {
					$RemovedChatMembers = @($MessageResult.eventDetail.members.id)
					ForEach ($RemovedChatMember in $RemovedChatMembers) {
						$ResponseJsonTemp = $null
						$RemovedMemberDetail = $null
						
						$ResponseJsonTemp = Invoke-GraphRequest -Uri "https://graph.microsoft.com/beta/users/$RemovedChatMember" -Headers $Headers
						$RemovedMemberDetail = $ResponseJsonTemp
						
						If ($RemovedMemberDetail -ne $null) {
							#Don't add duplicate removals
							If (($RemovalList | Where {$_.ChatId -eq $Chat.id -and $_.RemovedMemberId -eq $RemovedMemberDetail.id}).Count -eq 0) {
								$RemovalItem = New-Object -TypeName PSObject
								Add-Member -InputObject $RemovalItem -Name 'ChatId' -Value $Chat.id -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'ChatTopic' -Value $Chat.topic -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'ChatType' -Value $Chat.chatType -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'ChatLastMessageUTC' -Value '' -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'ChatMembers' -Value ($Chat.members.email -join "|") -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'RemovedMemberId' -Value $RemovedMemberDetail.id -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'RemovedMemberUpn' -Value $RemovedMemberDetail.userPrincipalName -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'RemovedMemberEmail' -Value $RemovedMemberDetail.mail -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'RemovedMemberEnabled' -Value $RemovedMemberDetail.accountEnabled -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'RemovedMemberReadded' -Value ($Chat.members.email -contains $RemovedMemberDetail.mail) -MemberType NoteProperty
								Add-Member -InputObject $RemovalItem -Name 'RemovedTimeUTC' -Value ([datetime]$MessageResult.createdDateTime).ToUniversalTime().ToString('MM-dd-yyyy HH:mm:ss') -MemberType NoteProperty
								
								$TempRemovalList = $TempRemovalList + $RemovalItem
								
								$f = $f + 1
								
								Write-Progress -Id 0 -Activity "Processing user ($a / $($UserList.Count)): $($User.UserPrincipalName)" -Status "Removals found under user: $f  |  Total Removals: $($RemovalList.Count + $TempRemovalList.Count)" -PercentComplete (($a / $UserList.Count) * 100)
							}
						}
						Else {
							Write-Host "-Error getting user: $RemovedChatMember. User deleted?"
						}
					}
					
					$Url = $ResponseJson.'@odata.nextLink'
				}
				ElseIf ([datetime]$MessageResult.createdDateTime -lt $StartDateTime) {
					$Url = $null
					#Break
				}
				Else {
					$Url = $ResponseJson.'@odata.nextLink'
				}
			}
			
			$d = $d + 1
		} While ($Url -ne $null)
		
		#Set last message time on each removal and add to array
		ForEach ($TempRemovalListItem in $TempRemovalList) {
			If ($LastChatMessageActivity.Year -ne 1) {
				$TempRemovalListItem.ChatLastMessageUTC = $LastChatMessageActivity.ToUniversalTime().ToString('MM-dd-yyyy HH:mm:ss')
			}
			$RemovalList = $RemovalList + $TempRemovalListItem
		}
		
		$c = $c + 1
	}
	
	Write-Host "-Total removal items: $f"
	
	Write-Progress -Id 2 -ParentId 1 -Activity 'Completed' -Completed
	Write-Progress -Id 3 -ParentId 2 -Activity 'Completed' -Completed
	
	$a = $a + 1
	Write-Host ""
}




$ErrorUsers | Out-File -FilePath $ErrorTxtPath
$RemovalList | Export-Csv -NoTypeInformation -Path $OutputCsvPath