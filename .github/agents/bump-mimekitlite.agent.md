---
name: Bump MimeKitLite Version
description: Automate MimeKitLite dependency version bumping with two-commit workflow, build validation, and PR creation
argument-hint: "Provide the new version for MimeKitLite (e.g., 4.15.0)"
---

# Bump MimeKitLite Version Agent

This GitHub Copilot agent automates the complete two-commit dependency bump workflow for MimeKitLite, including build validation and PR creation.

## Usage
```
@copilot /bump-mimekitlite newVersion=4.15.0
```

## Execution Flow

**Phase 1**: Validate input and gather current versions  
**Phase 2**: Create feature branch  
**Phase 3**: Commit 1 — Update MimeKitLite dependency across all files and validate build  
**Phase 4**: Commit 2 — Bump project version, update release notes, workflow, and changelog  
**Phase 5**: Push branch and create PR against `dev`

---

## Agent Instructions

### Phase 1: Gather Current Versions & Validate Input

1. **Validate the `newVersion` argument** — must match format `X.Y.Z` (e.g., `4.15.0`)
   - If not provided or invalid format, halt and report error

2. **Extract current MimeKitLite version** from `GmailAPIHelper/GmailAPIHelper.csproj`:
   - Search for pattern: `<PackageReference Include="MimeKitLite" Version="([^"]+)" />`
   - Store as `currentMimeVersion`

3. **Extract current project version** from `GmailAPIHelper/GmailAPIHelper.csproj`:
   - Search for pattern: `<Version>([^<]+)</Version>`
   - Store as `currentProjectVersion`

4. **Report findings to user**:
   ```
   Current versions detected:
     - MimeKitLite: [currentMimeVersion]
     - Project:     [currentProjectVersion]
   Bump target:
     - MimeKitLite: [newVersion] (new)
   ```

---

### Phase 2: Create Feature Branch

5. **Capture the existing Git identity before overriding it**:
   ```bash
   previousUserName=$(git config --get user.name)
   previousUserEmail=$(git config --get user.email)
   ```
   - Stores the original values so they can be restored after the PR workflow completes

6. **Configure git user for automated commits** (branch-specific):
   ```
   git config user.name "copilot-agent[bot]"
   git config user.email "copilot-agent@users.noreply.github.com"
   ```
   - Sets bot identity for commits on this feature branch only

7. **Check out `dev` branch first**:
   ```
   git checkout dev
   git pull origin dev
   ```
   - Ensures feature branch is based on latest `dev`

8. **Create branch name**: `bump/mimekitlite-X-Y-Z` (replace dots with dashes)

9. **Execute**: `git checkout -b [branchName]`
   - If branch already exists, delete it first: `git branch -D [branchName]`

---

### Phase 3: Commit 1 — Update Dependency Versions

10. **Update `GmailAPIHelper/GmailAPIHelper.csproj` using the .NET CLI**:
   - Execute: `dotnet add GmailAPIHelper/GmailAPIHelper.csproj package MimeKitLite --version [newVersion]`
   - This updates the existing `MimeKitLite` package reference in a way that works consistently in local environments and GitHub-hosted agents where the .NET SDK is available.
   - If the command fails, fall back to editing the project file directly and preserve the same `<PackageReference Include="MimeKitLite" Version="[newVersion]" />` format.

11. **Update the legacy .NET Framework test project with NuGet CLI**:
   - Execute: `nuget update GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj -Id MimeKitLite -Version [newVersion]`
   - This updates the package entry in `GmailAPIHelper.NET.Tests/packages.config` and refreshes the restored package assets without manually editing the old `packages.config` or `.csproj` files.
   - If `nuget` is unavailable, fall back to: `nuget install MimeKitLite -Version [newVersion] -OutputDirectory packages` followed by `nuget restore GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj`

12. **Refresh dependency state for the legacy test project**:
   - Execute: `nuget restore GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj`
   - If the project still needs its package graph rehydrated, execute: `msbuild GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj /t:Restore`
   - This keeps the update command-driven and avoids hand-editing `GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj` or any binding-redirect-related configuration.

13. **Validate build**:
    - Execute: `dotnet restore GmailAPIHelper.sln`
    - Execute: `dotnet build GmailAPIHelper.sln -c Release`
    - On build failure: rollback file changes (`git checkout .`), delete branch, report error and abort
    - On build success: continue to commit
    - If the agent is running locally, create the NuGet package by executing: `dotnet pack GmailAPIHelper/GmailAPIHelper.csproj -c Release`

14. **Create Commit 1**:
    - Stage files: `GmailAPIHelper/GmailAPIHelper.csproj`, `GmailAPIHelper.NET.Tests/packages.config`, `GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj`
    - Execute: `git add [files]`
    - Message: `MimeKitLite dependency update ('[currentMimeVersion]' -> '[newVersion]')`
    - Execute: `git commit -m "[message]"`

---

### Phase 4: Commit 2 — Bump Package Version & Update Release Files

15. **Calculate new project version**:
    - Parse `currentProjectVersion` (e.g., `1.12.1`) into parts: major, minor, patch
    - Increment patch version: `major.minor.(patch+1)`
    - Store as `newProjectVersion`

16. **Update `GmailAPIHelper/GmailAPIHelper.csproj`**:
    - Find: `<Version>[currentProjectVersion]</Version>`
    - Replace with: `<Version>[newProjectVersion]</Version>`
    - Find: `<PackageReleaseNotes>...</PackageReleaseNotes>`
    - Replace with: `<PackageReleaseNotes>1. MimeKitLite dependency update ('[currentMimeVersion]' -&gt; '[newVersion]').</PackageReleaseNotes>`

17. **Update `.github/workflows/publish-nuget-Package.yml`**:
    - Find: `NUGET_PACKAGE_NAME_VERSION: "GmailHelper.[currentProjectVersion].nupkg"`
    - Replace with: `NUGET_PACKAGE_NAME_VERSION: "GmailHelper.[newProjectVersion].nupkg"`

18. **Update `CHANGELOG.md`**:
    - Get today's date in Melbourne time (AEDT/AEST) using:
      ```
      $melbourneTime = [System.TimeZoneInfo]::ConvertTime([DateTime]::Now, [System.TimeZoneInfo]::FindSystemTimeZoneById('AUS Eastern Standard Time'))
      $todayDate = $melbourneTime.ToString('yyyy-MM-dd')
      ```
    - Prepend new entry at top of file under ## [Released] section:
    ```

    ## [newProjectVersion](https://www.nuget.org/packages/GmailHelper/[newProjectVersion]) - [today]
    ### Changed
    - MimeKitLite dependency update ('[currentMimeVersion]' -> '[newVersion]')

    ```

19. **Create Commit 2**:
    - Stage files: `GmailAPIHelper/GmailAPIHelper.csproj`, `.github/workflows/publish-nuget-Package.yml`, `CHANGELOG.md`
    - Execute: `git add [files]`
    - Message: `Nuget package creation - v[newProjectVersion]`
    - Execute: `git commit -m "[message]"`

---

### Phase 5: Push & Create PR

20. **Push feature branch**:
    - Execute: `git push origin [branchName]`
    - On push failure: report error and authentication/permission requirements

21. **Create pull request** using GitHub CLI:
    - Execute: `gh pr create --title "Bump **MimeKitLite**: [currentMimeVersion] → [newVersion]" --body "[body]" --base dev --head [branchName]`
    - PR body content:
    ```
    ## Dependency Update
    - **MimeKitLite**: [currentMimeVersion] → [newVersion]
    - **Project Version**: [currentProjectVersion] → [newProjectVersion]

    ## Changes
    - Updated dependency version
    - Bumped NuGet package version
    - Updated NuGet package release workflow configuration
    - Updated CHANGELOG

    ## Validation
    - [x] Build succeeds (Release configuration)
    - Tests will run in CI/CD on PR approval

    CC: @abhinavminhas
    ```

22. **Switch back to `dev` and restore the original Git identity**:
    - Execute: `git checkout dev`
    - Restore the previously captured identity using the same bash-style syntax:
      ```bash
      git config user.name "$previousUserName"
      git config user.email "$previousUserEmail"
      ```
      - If either previous value was empty, unset it with `git config --unset user.name` or `git config --unset user.email`

23. **Report success**:
    - Display PR URL
    - Show branch name
    - List all modified files
    - Show commit summaries

---

## Error Handling

- **Invalid version format**: Report expected format and halt
- **Build failure**: Rollback changes, delete branch, display error cause
- **Git/GitHub CLI not installed**: Report and suggest installation
- **Push/PR creation fails**: Report error with remediation steps
- **File not found**: Report missing file and halt

---

## Prerequisites

- Git configured: user.name=copilot-agent[bot], user.email=copilot-agent@users.noreply.github.com
- dotnet CLI installed and on PATH
- GitHub CLI (`gh`) installed for PR creation
- Working directory: Gmail-API-Helper repository root

---

## Example Execution

```
User request: @copilot /bump-mimekitlite newVersion=4.15.0

Agent output:
✓ Validating newVersion format: 4.15.0 ✓
✓ Current versions detected:
  - MimeKitLite: 4.14.0
  - Project:     1.12.1
✓ Creating branch: bump/mimekitlite-4-15-0
✓ Updating dependency versions in 3 files...
✓ Building solution (Release config)...
✓ Build successful
✓ Commit 1: MimeKitLite dependency update ('4.14.0' -> '4.15.0')
✓ Calculating new project version: 1.12.2
✓ Updating .csproj, workflow, and CHANGELOG...
✓ Commit 2: Nuget package creation - v1.12.2
✓ Pushing branch...
✓ Pull request created: https://github.com/abhinavminhas/Gmail-API-Helper/pull/XX
```

---

## Notes

- **Fully automated**: Copilot executes all steps without manual intervention
- **Build validation**: Release configuration only; integration tests run in CI/CD on PR
- **Version bump strategy**: Always increments project patch version (X.Y.Z → X.Y.Z+1)
- **PR target**: Always `dev` branch
- **Commit author**: copilot-agent[bot] (copilot-agent@users.noreply.github.com)
- **Reversible**: Branch can be deleted with `git branch -D [branchName]` to restart
