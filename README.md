# Report Automation for Microsoft Word

Report Automation is a native Windows Microsoft Word VSTO add-in written in C#/.NET Framework 4.8 and distributed with ClickOnce. It does not require Node.js, an HTTPS website, or `npm.cmd start` on user computers.

## Gender Swap

The **Home → Report Automation** ribbon group provides:

- **Gender Swap** — processes the entire document, including headers, footers, footnotes and other Word stories.
- **Swap Selection** — processes only the currently selected text.
- A match-count preview and confirmation before editing.
- One Word Undo record for the complete operation.
- Whole-word matching and preservation of lowercase, Title Case and UPPERCASE.

Rules:

- `he` → `her`, `she` → `him`
- `him` → `her`, `her` → `him`
- `his` → `hers`, `hers` → `his`
- `himself` ↔ `herself`
- `man` ↔ `woman`, `men` ↔ `women`

These direct rules intentionally follow the requested mappings. Review the result because English pronouns have different grammatical roles.

## Requirements

### End-user computers

- Windows
- Microsoft Word desktop
- .NET Framework 4.8
- Visual Studio 2010 Tools for Office Runtime

`setup.exe` checks and installs missing prerequisites when the computer can reach Microsoft's prerequisite download locations and the user has the required installation permissions.

### Developer/publisher computer

- Visual Studio 2022 with the **Office/SharePoint development** workload
- A code-signing certificate in `Cert:\CurrentUser\My`

The solution is [src/ReportAutomation.sln](src/ReportAutomation.sln).

## Create a local development certificate

Run once from PowerShell:

```powershell
.\scripts\New-DevelopmentCertificate.ps1
```

Copy the thumbprint printed by the script. This non-exportable self-signed certificate is for local testing only. Use an organisation-trusted Authenticode/code-signing certificate for production deployment.

## Publish a local ClickOnce installer

```powershell
.\scripts\Publish-ClickOnce.ps1 `
  -CertificateThumbprint "YOUR_CERTIFICATE_THUMBPRINT" `
  -Version "1.0.0.0"
```

Output is written to `artifacts\ClickOnce`. Install by running `setup.exe`. Keep `setup.exe`, `ReportAutomation.Vsto.vsto`, and the complete `Application Files` directory together.

## Publish to a network share

Create a versioned release in a stable UNC location:

```powershell
.\scripts\Publish-ClickOnce.ps1 `
  -CertificateThumbprint "YOUR_PRODUCTION_CERTIFICATE_THUMBPRINT" `
  -PublishPath "\\FileServer\Software\ReportAutomation" `
  -InstallUrl "\\FileServer\Software\ReportAutomation" `
  -Version "1.0.1.0"
```

Users run:

```text
\\FileServer\Software\ReportAutomation\setup.exe
```

For every release:

1. Increase the four-part `Version`.
2. Publish to the same `PublishPath` and `InstallUrl`.
3. Keep older version folders under `Application Files` so existing installations can update safely.
4. Sign with the same trusted production certificate, renewing it through your normal certificate-management process.

ClickOnce checks the configured installation location for updates when Word starts.

## Build and test

Build the add-in:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  .\src\ReportAutomation.sln /t:Rebuild /p:Configuration=Release `
  /p:SignManifests=true /p:ManifestCertificateThumbprint=YOUR_CERTIFICATE_THUMBPRINT
```

Build and run the dependency-free rule tests:

```powershell
& "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  .\tests\ReportAutomation.Rules.Tests\ReportAutomation.Rules.Tests.csproj `
  /t:Rebuild /p:Configuration=Release

.\tests\ReportAutomation.Rules.Tests\bin\Release\ReportAutomation.Rules.Tests.exe
```

## Security and operations

- Do not commit `.pfx`, `.snk`, private keys, or certificate passwords.
- Deploy the production certificate chain to Trusted Publishers and Trusted Root Certification Authorities through Group Policy or your device-management platform where appropriate.
- Pilot the installer with a small user group before broad deployment.
- ClickOnce is per-user. Use MSI/Intune instead if you later require a machine-wide installation.
