$ExpiresOn = Get-AutomationVariable -Name "ExpiresOn"
if ([DateTime]::Now -gt [DateTime]$ExpiresOn) {
  Write-Verbose "ExpiresOn has passed, getting new token"
  $RefreshToken = Get-AutomationVariable -Name 'RefreshToken'
  $RequestParameters = @{
    URI = "$(Get-AutomationVariable -Name 'AuthenticationURI')/refresh"
	Body = @{ token = $RefreshToken }
  }
  $Data = Invoke-RestMethod @RequestParameters
  Set-AutomationVariable -Name "AccessToken" -Value $Data.data.token.access_token
  $NewExpiresOn = [DateTime]::Now.AddSeconds($Data.data.token.expires_in)
  Set-AutomationVariable -Name "ExpiresOn" -Value $NewExpiresOn
}