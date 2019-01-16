# Scripts

These scripts generate the NSA Cybersecurity library zip file and the [publications.md](../publications page):
* [CybersecurityLibrary.psm1](./CybersecurityLibrary.psm1) is a PowerShell module that contains logic for parsing nsa.gov pages that contain NSA Cybersecurity publications.
* [GenerateCybersecurityLibrary.ps1](./GenerateCybersecurityLibrary.ps1) calls the PowerShell module to generate the zip file and the publications page.