---
name: Bump MimeKitLite Version
description: Automate MimeKitLite dependency version bumping with two-commit workflow, build validation, and PR creation
argument-hint: "Provide the new version for MimeKitLite (e.g., 4.16.0)"
---

# Bump MimeKitLite Version Agent

This GitHub Copilot agent automates the complete two-commit dependency bump workflow for MimeKitLite, including build validation and PR creation.

## Usage
```
@copilot /bump-mimekit-lite newVersion=4.15.0
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

1. **Validate the `newVersion` argument** — must match format `X.Y.Z` (e.g., `4.16.0`)
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

5. **Check out `dev` branch first**:
   ```
   git checkout dev
   git pull origin dev
   ```
   - Ensures feature branch is based on latest `dev`

6. **Create branch name**: `bump/mimekit-lite-X-Y-Z` (replace dots with dashes)

7. **Execute**: `git checkout -b [branchName]`
   - If branch already exists, delete it first: `git branch -D [branchName]`

---

### Phase 3: Commit 1 — Update Dependency Versions

7. **Update `GmailAPIHelper/GmailAPIHelper.csproj`**:
   - Find: `<PackageReference Include="MimeKitLite" Version="[currentMimeVersion]" />`
   - Replace with: `<PackageReference Include="MimeKitLite" Version="[newVersion]" />`

8. **Update `GmailAPIHelper.NET.Tests/packages.config`**:
   - Find: `<package id="MimeKitLite" version="[currentMimeVersion]" targetFramework=`
   - Replace with: `<package id="MimeKitLite" version="[newVersion]" targetFramework=`

9. **Validate build**:
    - Execute: `dotnet restore GmailAPIHelper.sln`
    - Execute: `dotnet build GmailAPIHelper.sln -c Release`
    - On build failure: rollback file changes (`git checkout .`), delete branch, report error and abort
    - On build success: continue to commit

10. **Create Commit 1**:
    - Stage files: `GmailAPIHelper/GmailAPIHelper.csproj`, `GmailAPIHelper.NET.Tests/packages.config`
    - Execute: `git add [files]`
    - Message: `MimeKitLite dependency update ('[currentMimeVersion]' -> '[newVersion]')`
    - Execute: `git commit -m "[message]"`

---

### Phase 4: Commit 2 — Bump Package Version & Update Release Files

11. **Calculate new project version**:
    - Parse `currentProjectVersion` (e.g., `1.12.2`) into parts: major, minor, patch
    - Increment patch version: `major.minor.(patch+1)`
    - Store as `newProjectVersion`

12. **Update `GmailAPIHelper/GmailAPIHelper.csproj`**:
    - Find: `<Version>[currentProjectVersion]</Version>`
    - Replace with: `<Version>[newProjectVersion]</Version>`
    - Find: `<PackageReleaseNotes>...</PackageReleaseNotes>`
    - Replace with: `<PackageReleaseNotes>1. MimeKitLite dependency update ('[currentMimeVersion]' -> '[newVersion]').</PackageReleaseNotes>`

13. **Update `.github/workflows/publish-nuget-Package.yml`**:
    - Find: `NUGET_PACKAGE_NAME_VERSION: "GmailHelper.[currentProjectVersion].nupkg"`
    - Replace with: `NUGET_PACKAGE_NAME_VERSION: "GmailHelper.[newProjectVersion].nupkg"`

14. **Update `CHANGELOG.md`**:
    - Get today's date in Melbourne time (AEDT/AEST) using:
      ```
      $melbourneTime = [System.TimeZoneInfo]::ConvertTime([DateTime]::Now, [System.TimeZoneInfo]::FindSystemTimeZoneById('AUS Eastern Standard Time'))
      $todayDate = $melbourneTime.ToString('yyyy-MM-dd')
      ```
    - Prepend new entry at top of file:
    ```
    ## [newProjectVersion] - [today]
    ### Changed
    - MimeKitLite dependency update ('[currentMimeVersion]' -> '[newVersion]')

    ```

15. **Create Commit 2**:
    - Stage files: `GmailAPIHelper/GmailAPIHelper.csproj`, `.github/workflows/publish-nuget-Package.yml`, `CHANGELOG.md`
    - Execute: `git add [files]`
    - Message: `Nuget package creation - v[newProjectVersion]`
    - Execute: `git commit -m "[message]"`

---

### Phase 5: Push & Create PR

16. **Push feature branch**:
    - Execute: `git push origin [branchName]`
    - On push failure: report error and authentication/permission requirements

17. **Create pull request** using GitHub CLI:
    - Execute: `gh pr create --title "Nuget Package Creation - v[newProjectVersion]" --body "[body]" --base dev --head [branchName]`
    - PR body content:
    ```
    ## Dependency Update
    - **MimeKitLite**: [currentMimeVersion] → [newVersion]
    - **Project Version**: [currentProjectVersion] → [newProjectVersion]

    ## Changes
    - Updated dependency version across 2 project files
    - Bumped NuGet package version
    - Updated workflow configuration
    - Updated CHANGELOG

    ## Validation
    - [x] Build succeeds (Release configuration)
    - Tests will run in CI/CD on PR approval

    CC: @abhinavminhas
    ```

18. **Report success**:
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

- Git configured: user.name=abhinavminhas, user.email=abhinavminhas@users.noreply.github.com
- dotnet CLI installed and on PATH
- GitHub CLI (`gh`) installed for PR creation
- Working directory: Gmail-API-Helper repository root

---

## Example Execution

```
User request: @copilot /bump-mimekit-lite newVersion=4.16.0

Agent output:
✓ Validating newVersion format: 4.16.0 ✓
✓ Current versions detected:
  - MimeKitLite: 4.15.0
  - Project:     1.12.2
✓ Creating branch: bump/mimekit-lite-4-16-0
✓ Updating dependency versions in 2 files...
✓ Building solution (Release config)...
✓ Build successful
✓ Commit 1: MimeKitLite dependency update ('4.15.0' -> '4.16.0')
✓ Calculating new project version: 1.12.3
✓ Updating .csproj, workflow, and CHANGELOG...
✓ Commit 2: Nuget package creation - v1.12.3
✓ Pushing branch...
✓ Pull request created: https://github.com/abhinavminhas/Gmail-API-Helper/pull/XX
```

---

## Notes

- **Fully automated**: Copilot executes all steps without manual intervention
- **Build validation**: Release configuration only; integration tests run in CI/CD on PR
- **Version bump strategy**: Always increments project patch version (X.Y.Z → X.Y.Z+1)
- **PR target**: Always `dev` branch
- **Commit author**: abhinavminhas (abhinavminhas@users.noreply.github.com)
- **Fewer files than Gmail API agent**: MimeKitLite updates only 2 files in Commit 1
- **Reversible**: Branch can be deleted with `git branch -D [branchName]` to restart
