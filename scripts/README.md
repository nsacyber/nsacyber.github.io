# Scripts

These scripts generate the NSA Cybersecurity library zip file and the [publications page](./../publications.md):
* [CybersecurityLibrary.psm1](./CybersecurityLibrary.psm1) is a PowerShell module that contains logic for parsing nsa.gov pages that contain NSA Cybersecurity publications.
* [GenerateCybersecurityLibrary.ps1](./GenerateCybersecurityLibrary.ps1) calls the PowerShell module to generate the zip file and the publications page.

To execute the script:
1. Open a PowerShell prompt as user
1. Change directory to the folder that contains both scripts
1. Type `. .\GenerateCybersecurityLibrary.ps1`, press Enter, and wait for the script to complete