$newVersion = (Get-Date).toString("yyyy-MM-dd@hh-mm")

$a = Get-Content dist\version.json -raw | ConvertFrom-Json
$a.build = $newVersion
$a | ConvertTo-Json  | set-content dist\version.json