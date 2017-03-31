<h3>Description</h3>

These scripts will allow users to unlock files on share drives without admin rights on the server.

<h3>Installation</h3>

<h4>Client</h4>
Download UnlockRequest.vbs and place it on your shared drive. S:\UnlockRequest is used in this example.

Edit UnlockRequest.vbs to change paths and e-mail info

Change permissions on UnlockRequest.vbs so that Users only have Read&Execute rights.

(optional) Create an empty file named UnlockRequest.txt. Anyone who has write permissions to this file will be able to request unlocks. Change permissions accordingly.



<h4>Server</h4>

Download https://live.sysinternals.com/psfile64.exe and place under C:\Scripts on the file server



Download UnlockRequest.ps1 & RunUnlockRequest.ps1 and place under C:\Scripts on the file server

Edit UnlockRequest.ps1 to change paths

Configure RunUnlockRequest.ps1 to run as scheduled task or as a windows service. SYSTEM account will be sufficient.
