# Installer

> **Note** — this directory was added on the user's request.
> SPEC §15 originally declared that this project would not
> ship an installer, but the user is now distributing the
> add-in to non-developer teammates, so we ship a thin
> PowerShell installer. A full MSI / EXE builder
> (`office-addin-installer`) is out of scope here — if you
> need one for enterprise Group Policy rollout, see the
> "Going further" section below.

## What it does

Copies the project into a per-user Office add-in manifest
directory so Word, Excel, and PowerPoint can sideload it
under **My Add-ins**:

```
%LOCALAPPDATA%\Microsoft\Office\16.0\WEF\hopeOffice\
├── manifest.xml
├── public\        # static assets (icons, etc.)
└── src\           # dev source — Office reads it through
                   # the dev server (https://localhost:3721)
```

After the install, the user restarts Office and goes to
**Insert → My Add-ins → hopeOffice** to mount it.

The install does **not** require admin rights — everything
lands in the per-user `%LOCALAPPDATA%` tree.

## Usage

From the project root:

```powershell
# Dry-run — print what would happen without touching anything
pwsh installer/install.ps1 -DryRun

# Dev install (copies src/, requires `npm run dev` later)
pwsh installer/install.ps1

# Dev install + register the install dir as a trusted catalog
# (skips the "do you trust this?" prompt in Office)
pwsh installer/install.ps1 -TrustCatalog

# Production install (copies dist/, requires a prior `npm run build`
# and a way to serve the built assets over HTTPS)
pwsh installer/install.ps1 -Production

# Uninstall — removes the install dir and any trusted-catalog entry
pwsh installer/install.ps1 -Uninstall
```

After install, **restart Word / Excel / PowerPoint** so the
manifest is picked up.

## Going further: MSI / EXE

For enterprise distribution, Microsoft's official
[`office-addin-installer`](https://www.npmjs.com/package/office-addin-installer)
package builds MSI / EXE installers configured for Group
Policy. The two prerequisites are a `.pfx` code-signing
certificate and an `installer/manifest.xml` that points at
the production SourceUrl. We don't ship that configuration
here because it's tightly coupled to your organization's
network topology and signing infrastructure.
