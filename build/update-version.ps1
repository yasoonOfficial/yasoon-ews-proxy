$newVersion = (Get-Date).toString("yyyy-MM-dd@hh-mm")

$a = Get-Content version.json -raw | ConvertFrom-Json
$a.build = $newVersion
$a | ConvertTo-Json  | set-content version.json