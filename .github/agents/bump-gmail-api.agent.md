---
name: Bump Gmail API Version
description: Automate Gmail API dependency version bumping with two-commit workflow, build validation, and PR creation
argument-hint: "Provide the new version for Google.Apis.Gmail.v1 (e.g., 1.73.0.4029)"
---

# Bump Gmail API Version Agent

This GitHub Copilot agent automates the complete two-commit dependency bump workflow for Gmail API, including build validation and PR creation.

## Usage
```
@copilot /bump-gmail-api newVersion=1.73.0.4029
```

## Execution Flow

**Phase 1**: Validate input and gather current versions  
**Phase 2**: Create feature branch  
**Phase 3**: Commit 1 — Update Gmail API dependency across all files and validate build  
**Phase 4**: Commit 2 — Bump project version, update release notes, workflow, and changelog  
**Phase 5**: Push branch and create PR against `dev`

---

## Agent Instructions

### Phase 1: Gather Current Versions & Validate Input

1. **Validate the `newVersion` argument** — must match format `X.Y.Z.W` (e.g., `1.73.0.4029`)
   - If not provided or invalid format, halt and report error

2. **Extract current Gmail API version** from `GmailAPIHelper/GmailAPIHelper.csproj`:
   - Search for pattern: `<PackageReference Include="Google.Apis.Gmail.v1" Version="([^"]+)" />`
   - Store as `currentGmailVersion`

3. **Extract current project version** from `GmailAPIHelper/GmailAPIHelper.csproj`:
   - Search for pattern: `<Version>([^<]+)</Version>`
   - Store as `currentProjectVersion`

4. **Report findings to user**:
   ```
   Current versions detected:
     - Gmail API: [currentGmailVersion]
     - Project:   [currentProjectVersion]
   Bump target:
     - Gmail API: [newVersion] (new)
   ```

---

### Phase 2: Create Feature Branch

5. **Configure git user for automated commits** (branch-specific):
   ```
   git config user.name "copilot-agent[bot]"
   git config user.email "copilot-agent@users.noreply.github.com"
   ```
   - Sets bot identity for commits on this feature branch only

6. **Check out `dev` branch first**:
   ```
   git checkout dev
   git pull origin dev
   ```
   - Ensures feature branch is based on latest `dev`

7. **Create branch name**: `bump/gmail-api-X-Y-Z-W` (replace dots with dashes)

8. **Execute**: `git checkout -b [branchName]`
   - If branch already exists, delete it first: `git branch -D [branchName]`

---

### Phase 3: Commit 1 — Update Dependency Versions

9. **Update `GmailAPIHelper/GmailAPIHelper.csproj`**:
   - Find: `<PackageReference Include="Google.Apis.Gmail.v1" Version="[currentGmailVersion]" />`
   - Replace with: `<PackageReference Include="Google.Apis.Gmail.v1" Version="[newVersion]" />`

10. **Update `GmailAPIHelper.NET.Tests/packages.config`**:
   - Extract base version from `[newVersion]` (remove last component): `[baseVersion] = [newVersion]` minus last `.W`
   - Update Gmail API package:
     - Find: `<package id="Google.Apis.Gmail.v1" version="[currentGmailVersion]" targetFramework=`
     - Replace with: `<package id="Google.Apis.Gmail.v1" version="[newVersion]" targetFramework=`
   - Update related Google.Apis packages to base version:
     - Find: `<package id="Google.Apis" version="[currentBaseVersion]" targetFramework=`
     - Replace with: `<package id="Google.Apis" version="[baseVersion]" targetFramework=`
     - Find: `<package id="Google.Apis.Auth" version="[currentBaseVersion]" targetFramework=`
     - Replace with: `<package id="Google.Apis.Auth" version="[baseVersion]" targetFramework=`
     - Find: `<package id="Google.Apis.Core" version="[currentBaseVersion]" targetFramework=`
     - Replace with: `<package id="Google.Apis.Core" version="[baseVersion]" targetFramework=`

11. **Update `GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj`** (if binding redirects exist):
   - Find any hardcoded version references for Google.Apis.Gmail.v1
   - Replace all occurrences with `[newVersion]`

12. **Validate build**:
    - Execute: `dotnet restore GmailAPIHelper.sln`
    - Execute: `dotnet build GmailAPIHelper.sln -c Release`
    - On build failure: rollback file changes (`git checkout .`), delete branch, report error and abort
    - On build success: continue to commit

13. **Create Commit 1**:
    - Stage files: `GmailAPIHelper/GmailAPIHelper.csproj`, `GmailAPIHelper.NET.Tests/packages.config`, `GmailAPIHelper.NET.Tests/GmailAPIHelper.NET.Tests.csproj`
    - Execute: `git add [files]`
    - Message: `Gmail API dependency update ('[currentGmailVersion]' -> '[newVersion]')`
    - Execute: `git commit -m "[message]"`

---

### Phase 4: Commit 2 — Bump Package Version & Update Release Files

14. **Calculate new project version**:
    - Parse `currentProjectVersion` (e.g., `1.12.1`) into parts: major, minor, patch
    - Increment patch version: `major.minor.(patch+1)`
    - Store as `newProjectVersion`

15. **Update `GmailAPIHelper/GmailAPIHelper.csproj`**:
    - Find: `<Version>[currentProjectVersion]</Version>`
    - Replace with: `<Version>[newProjectVersion]</Version>`
    - Find: `<PackageReleaseNotes>...</PackageReleaseNotes>`
    - Replace with: `<PackageReleaseNotes>1. Gmail API dependency update ('[currentGmailVersion]' -&gt; '[newVersion]').</PackageReleaseNotes>`

16. **Update `.github/workflows/publish-nuget-Package.yml`**:
    - Find: `NUGET_PACKAGE_NAME_VERSION: "GmailHelper.[currentProjectVersion].nupkg"`
    - Replace with: `NUGET_PACKAGE_NAME_VERSION: "GmailHelper.[newProjectVersion].nupkg"`

17. **Update `CHANGELOG.md`**:
    - Get today's date in Melbourne time (AEDT/AEST) using:
      ```
      $melbourneTime = [System.TimeZoneInfo]::ConvertTime([DateTime]::Now, [System.TimeZoneInfo]::FindSystemTimeZoneById('AUS Eastern Standard Time'))
      $todayDate = $melbourneTime.ToString('yyyy-MM-dd')
      ```
    - Prepend new entry at top of file under ## [Released] section:
    ```

    ## [newProjectVersion](https://www.nuget.org/packages/GmailHelper/[newProjectVersion]) - [today]
    ### Changed
    - Gmail API dependency update ('[currentGmailVersion]' -> '[newVersion]')

    ```

18. **Create Commit 2**:
    - Stage files: `GmailAPIHelper/GmailAPIHelper.csproj`, `.github/workflows/publish-nuget-Package.yml`, `CHANGELOG.md`
    - Execute: `git add [files]`
    - Message: `Nuget package creation - v[newProjectVersion]`
    - Execute: `git commit -m "[message]"`

---

### Phase 5: Push & Create PR

19. **Push feature branch**:
    - Execute: `git push origin [branchName]`
    - On push failure: report error and authentication/permission requirements

20. **Create pull request** using GitHub CLI:
    - Execute: `gh pr create --title "Nuget Package Creation - v[newProjectVersion]" --body "[body]" --base dev --head [branchName]`
    - PR body content:
    ```
    ## Dependency Update
    - **Gmail API**: [currentGmailVersion] → [newVersion]
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

21. **Report success**:
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
User request: @copilot /bump-gmail-api newVersion=1.73.0.4029

Agent output:
✓ Validating newVersion format: 1.73.0.4029 ✓
✓ Current versions detected:
  - Gmail API: 1.73.0.3987
  - Project:   1.12.1
✓ Creating branch: bump/gmail-api-1-73-0-4029
✓ Updating dependency versions in 3 files...
✓ Building solution (Release config)...
✓ Build successful
✓ Commit 1: Gmail API dependency update ('1.73.0.3987' -> '1.73.0.4029')
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
