$siteID = 'mmoeller.sharepoint.com,0f135c4b-ca75-49bf-b019-b790701581da,d68ebd68-1318-4b51-987b-5190155a833e'
$listID = '16fcdef9-65e3-4786-8ff9-a385f7444845'
$tenant = "mmoeller"
Connect-PnPOnline -Url "https://$tenant-admin.sharepoint.com"
$config = @{}
$config.siteID = $siteID
$config.listID = $listID
$json = $config | ConvertTo-JSON -Depth 2
Set-PnPStorageEntity -Key "DocReviewConfig" -Value $json.ToString() -Comment "Config DocReview" -Description "DocReview Teams solution Config"