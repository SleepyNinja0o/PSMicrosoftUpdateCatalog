# Microsoft Update Catalog PowerShell Module
Simple PowerShell module for querying the Microsoft Update Catalog and pulling update details + download links
<br/>
<br/>
## Examples
### Retrieve First 25 Catalog Search Results
```Get-MicrosoftUpdates -SearchText "KB5035307"```
<br/>
<br/>
### Retrieve Next Page of 25 Search Results (Ran after command above)
```Get-MicrosoftUpdates -NextPage```
<br/>
<br/>
### Retrieve Download Links for an Update
```Get-MicrosoftUpdateDownload -UpdateID '7e48a2bb-067c-4634-bf89-0038cf5abd0b'```
<br/>
<br/>
### Retrieve Details About an Update
```Get-MicrosoftUpdate -UpdateID '7e48a2bb-067c-4634-bf89-0038cf5abd0b'```