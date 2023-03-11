Param(
	[Parameter(Mandatory=$true)]
	[string] $CsvPath,
	
	[Parameter(Mandatory=$true)]
	[string] $TenantId,
	
	[Parameter(Mandatory=$true)]
	[string] $AppId,
	
	[Parameter(Mandatory=$true)]
	[string] $AppSecret
)

Import-Module MSAL.PS

$RemovalInstanceList = Import-Csv -Path $CsvPath




Function Connect-TeamsAPI {
	If ($Token.ExpiresOn.LocalDateTime -gt (Get-Date).AddMinutes(15)) {
		Return
	}
	Else {
		Write-Host 'Refreshing Teams API access token...'
	}
	
	$AppSecretString = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force
	$global:Token = Get-MsalToken -ClientId $AppId -ClientSecret $AppSecretString -TenantId $TenantId -ForceRefresh
	$global:Headers = @{Authorization = "Bearer $($Token.AccessToken)"}
}




ForEach ($RemovalInstance in $RemovalInstanceList) {
	Connect-TeamsAPI
	
	$Body = @{
		"@odata.type" = "#microsoft.graph.aadUserConversationMember"
		"User@odata.bind" = "https://graph.microsoft.com/beta/users/$($RemovalInstance.RemovedMemberId)"
		VisibleHistoryStartDateTime = "0001-01-01T00:00:00Z"
		Roles = @(
			"owner"
		)
	}
	
	$Body = $Body | ConvertTo-Json
	
	$Response = Invoke-WebRequest -Uri "https://graph.microsoft.com/beta/chats/$($RemovalInstance.ChatId)/members" -Headers $Headers -ContentType 'application/json' -Method POST -Body $Body -Verbose
	$ResponseJson = ($Response | ConvertFrom-Json)
}