Connect-PnPOnline -Url '<SiteURL>' -Credentials -UseWebLogin
$targetList = Get-PnPList -Identity "<List Title>"
$targetField = Get-PnPField -List $targetList -Identity "<Field Internal Name>"
$targetField.ClientSideComponentId = "<FieldCustomizer ID>"
$targetField.Update()
Invoke-PnPQuery