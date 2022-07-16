[CmdletBinding()]
param(
  [Parameter()]
  [Switch]
  $Force,

  [Parameter()]
  [Switch]
  # Used to notify the user of the change
  $Notify,

  [Parameter()]
  [String[]]
  $MailTo = @("nat@natfan.io","nathanwindisch@gmail.com"),

  [Parameter()]
  [String]
  $MailFrom = "spotify@sys.wnd.sh",

  [Parameter()]
  [String]
  $MailServer = "smtp.natfan.io"
)
$StartOfWeek = [DateTime]::Today; while($StartOfWeek.DayOfWeek -ne "Monday"){$StartOfWeek = $StartOfWeek.AddDays(-1)} # Just a quick Start of Week calculation
$EndOfWeek   = $StartOfWeek.AddDays(7)
$StartOfWeek = $StartOfWeek.ToShortDateString()
$EndOfWeek   = $EndOfWeek.ToShortDateString()
$CurrentDiscoverWeeklyName = "DiscoverWeekly_$StartOfWeek"

Write-Verbose "Start of Week: $StartOfWeek"
Write-Verbose "End of Week: $EndOfWeek"
Write-Verbose "Current Discover Weekly name: $CurrentDiscoverWeeklyName"

$null = .\Connect-Spotify.ps1

$AccessToken = Get-AutomationVariable -Name 'AccessToken'

###

Write-Verbose "Starting Spotify API calls"
$MyPlaylists = [Collections.ArrayList]@()
$BaseURI = "https://api.spotify.com/v1"
$Headers = @{ Authorization = "Bearer $AccessToken" }
$PSDefaultParameterValues["Invoke-RestMethod:Headers"] = $Headers
$MyID = $(Invoke-RestMethod -URI "$BaseURI/me").id
Write-Verbose "User ID: $MyID"

$GetMyPlaylistsURI = "$BaseURI/me/playlists?limit=50"
$GetMyPlaylistsOffset = 0
$TempPlaylists = Invoke-RestMethod -URI $GetMyPlaylistsURI
$MaxPlaylists = $TempPlaylists.total
Write-Verbose "Maximum playlists: $MaxPlaylists"

$TempPlaylists.Items.ForEach({$null = $MyPlaylists.Add($_)})
while($MyPlaylists.Count -lt $MaxPlaylists) {
  $GetMyPlaylistsOffset = $GetMyPlaylistsOffset + 50
  $GetMyPlaylistsURI = "$GetMyPlaylistsURI&offset=$GetMyPlaylistsOffset"
  $TempPlaylists = Invoke-RestMethod -URI $GetMyPlaylistsURI
  $TempPlaylists.Items.ForEach({$null = $MyPlaylists.Add($_)})
}

$DiscoverWeekly = $MyPlaylists.Where({$_.Name -eq "Discover Weekly"})
if (-NOT $DiscoverWeekly) {
  Write-Warning "No Discover Weekly playlist found. This shouldn't happen to be honest..."
  throw "Discover Weekly playlist not found, exiting"
}


$CurrentDiscoverWeekly = $MyPlaylists.Where({$_.Name -eq $CurrentDiscoverWeeklyName})
if (-NOT $CurrentDiscoverWeekly) {
  Write-Verbose "No Current Discover Weekly playlist exists"
  $CreatePlaylistParameters = @{
    Method = "POST"
    URI = "$BaseURI/users/$MyID/playlists"
    Body = @{
      name = $CurrentDiscoverWeeklyName
      description = "Discover Weeklies from $StartOfWeek to $EndOfWeek"
      public = $false
    } | ConvertTo-JSON
  }
  Write-Verbose "No Current Discover Weekly playlist exists"
  Write-Verbose "Attempting to create $CurrentDiscoverWeeklyName"
  try {
    $CurrentDiscoverWeekly = Invoke-RestMethod @CreatePlaylistParameters
    Write-Verbose "Successfully created $CurrentDiscoverWeeklyName"
  } catch {
    Write-Verbose "Failed to create new Discover Weekly playlist [$CurrentDiscoverWeeklyName]: $_" -Level Critical
    throw "Failed to create new Discover Weekly playlist"
  }
}

if (($CurrentDiscoverWeekly.Tracks.Total -eq 30) -and (-NOT $Force)) {
  Write-Warning "$CurrentDiscoverWeeklyName already has 30 songs. If you want to wipe and re-copy this playlist, please re-run this command with the -Force flag"
  throw "$CurrentDiscoverWeeklyName already has 30 songs. If you want to wipe and re-copy this playlist, please re-run this command with the -Force flag"
}

if ($CurrentDiscoverWeekly.Tracks.Total -ne 0 -and $CurrentDiscoverWeekly.Tracks.Total -lt 30) {
  Write-Verbose "Total tracks is between 1 and 29, which means that it was populated before."
  if ($Automated) {
    Write-Warning "Running in automation mode, unable to prompt for choice. Defaulting to exiting"
    exit
  }
}

if ($CurrentDiscoverWeekly.Tracks.Total -gt 0) {
  Write-Verbose "Current Discover Weekly is invalid and needs to be recreated"
  Write-Verbose "Querying what tracks to delete from the playlist"
  $TracksToRemove = Invoke-RestMethod -URI "$BaseURI/playlists/$($CurrentDiscoverWeekly.ID)/tracks"
  Write-Verbose "Found $($TracksToWipe.Tracks.Total) amount of tracks to remove"
  $TracksToRemoveX = $TracksToRemove.Items.Track | Select-Object @{N="uri";E={"spotify:track:$($_.ID)"}}
  Write-Verbose "Removing the following tracks: $($TracksToRemove.Items.Track.ID | ConvertTo-JSON -Compress)"
  $RemoveTracksParameters = @{
    Method = "DELETE"
    URI = "$BaseURI/playlists/$($CurrentDiscoverWeekly.ID)/tracks"
    Body = ConvertTo-JSON @{
      tracks = @($TracksToRemoveX)
    }
  }
  Write-Verbose "Attempting to remove tracks from playlist [$CurrentDiscoverWeeklyName]"
  try {
    $null = Invoke-RestMethod @RemoveTracksParameters
    Write-Verbose "Successfully removed tracks from playlist [$CurrentDiscoverWeeklyName]"
  } catch {
    Write-Warning "Failed to remove tracks from playlist [$CurrentDiscoverWeeklyName]: $_"
  }
}

Write-Verbose "Copying Discover Weekly tracks to $CurrentDiscoverWeeklyName"
$TracksToAdd = Invoke-RestMethod -URI "$BaseURI/playlists/$($DiscoverWeekly.id)/tracks"
Write-Verbose "Found $($TracksToAdd.Tracks.Total) amount of tracks to add"
Write-Verbose "Adding the following tracks: $($TracksToAdd.Items.Track.ID | ConvertTo-JSON -Compress)"
$TracksToAddX = $TracksToAdd.Items.Track.ForEach({return "spotify:track:$($_.ID)"}) -join ","
$AddTracksParameters = @{
  Method = "POST"
  URI = "$BaseURI/playlists/$($CurrentDiscoverWeekly.ID)/tracks?uris=$TracksToAddX"
  Body = ConvertTo-JSON @{
    tracks = @($TracksToAddX)
  }
}
Write-Verbose "Attempting to add tracks to playlist [$CurrentDiscoverWeeklyName]"
try {
  $null = Invoke-RestMethod @AddTracksParameters
  Write-Verbose "Successfully added tracks to playlist [$CurrentDiscoverWeeklyName]"
} catch {
  Write-Warning "Failed to add tracks to playlist [$CurrentDiscoverWeeklyName]: $_"
}

if ($Notify) {
  $OpenSpotifyLink = "https://open.spotify.com/playlist/$($CurrentDiscoverWeekly.ID)"
  # Set up the email notifications
  $SendMailMessageParameters = @{
    To = "blackhole@natfan.io"
    CC = $MailTo
    From = $MailFrom
    Subject = "[SpotifyDiscoverWeekly] Backup finished!"
    Body = "Your weekly backup of Spotify's Discover Weekly has finished. It's called '$CurrentDiscoverWeeklyName' and you can access it by this link:<br /><a href='$OpenSpotifyLink'>$OpenSpotifyLink</a>"
    BodyAsHTML = $true
    SMTPServer = $MailServer
  }
  Send-MailMessage @SendMailMessageParameters
}