# Contributing to OfficeScrubC2R

Thank you for your interest in contributing to OfficeScrubC2R! This document provides guidelines and instructions for contributing.

## Code of Conduct

- Be respectful and inclusive
- Provide constructive feedback
- Focus on what is best for the community

## How Can I Contribute?

### Reporting Bugs

Before creating bug reports, please check existing issues. When creating a bug report, include:

- **Clear title** and description
- **Steps to reproduce** the problem
- **Expected behavior** vs actual behavior
- **PowerShell version**: `$PSVersionTable`
- **Windows version**: `winver`
- **Log files**: From `$env:TEMP\OfficeScrubC2R\`
- **Screenshots** if applicable

### Suggesting Enhancements

Enhancement suggestions are tracked as GitHub issues. Include:

- **Clear title** and description
- **Use case** - why is this enhancement useful?
- **Current workaround** if any exists
- **Examples** of how it would work

### Pull Requests

1. **Fork** the repository
2. **Create a branch**: `git checkout -b feature/amazing-feature`
3. **Make your changes** following the coding standards
4. **Test thoroughly** on Windows PowerShell 5.1 and PowerShell 7+
5. **Commit** with clear messages: `git commit -m 'Add amazing feature'`
6. **Push** to your fork: `git push origin feature/amazing-feature`
7. **Open a Pull Request** with a clear description

## Development Setup

### Prerequisites

```powershell
# Install required modules
Install-Module -Name PSScriptAnalyzer -Scope CurrentUser
Install-Module -Name Pester -Scope CurrentUser

# Clone the repository
git clone https://github.com/Calvindd2f/OfficeScrubC2R.git
cd OfficeScrubC2R
```

### Building

```powershell
# Build the native DLL
.\build.ps1

# Import the module
Import-Module .\OfficeScrubC2R.psd1 -Force

# Run tests (when available)
Invoke-Pester
```

## Coding Standards

### PowerShell

- Use **PascalCase** for function names
- Use **Approved Verbs**: `Get-Verb`
- Add **comment-based help** for all functions
- Use **`[CmdletBinding()]`** for advanced functions
- Follow **PowerShell best practices**
- Use **explicit parameter types**

Example:

```powershell
function Get-Something {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Name,

        [switch]$Force
    )

    # Implementation
}
```

### C#

- Follow **.NET naming conventions**
- Add **XML documentation comments**
- Use **explicit types**
- Handle **errors gracefully**
- Consider **performance implications**

Example:

```csharp
/// <summary>
/// Does something useful
/// </summary>
/// <param name="value">The value to process</param>
/// <returns>The processed result</returns>
public string DoSomething(string value)
{
    if (string.IsNullOrEmpty(value))
        throw new ArgumentNullException(nameof(value));

    // Implementation
}
```

## Testing

### Manual Testing

```powershell
# Test detection only (safe)
Invoke-OfficeScrubC2R -DetectOnly -Verbose

# Test on a VM (recommended for full testing)
# Use Windows Sandbox or Hyper-V
```

### Automated Testing

```powershell
# Run PSScriptAnalyzer
Invoke-ScriptAnalyzer -Path . -Recurse

# Run Pester tests
Invoke-Pester -Path .\tests\
```

### Test Checklist

- [ ] Works on Windows PowerShell 5.1
- [ ] Works on PowerShell 7+
- [ ] Works on Windows 10/11
- [ ] Handles errors gracefully
- [ ] Logs important operations
- [ ] Doesn't break existing functionality
- [ ] Passes PSScriptAnalyzer
- [ ] Documentation updated

## Documentation

- Update **README.md** for user-facing changes
- Update **CHANGELOG.md** with your changes
- Update **help documentation** if adding/changing functions
- Add **inline comments** for complex logic
- Update **BUILD.md** if changing build process

## Commit Messages

Use clear, descriptive commit messages:

```
Add feature to detect Office 2021

- Implement detection for Office 2021 C2R
- Add registry key checks for Office 2021
- Update tests to cover Office 2021

Fixes #123
```

Format:

- **First line**: Short summary (50 chars max)
- **Body**: Detailed explanation (wrap at 72 chars)
- **Footer**: Reference issues

## Version Numbering

We follow [Semantic Versioning](https://semver.org/):

- **MAJOR**: Breaking changes
- **MINOR**: New features (backward compatible)
- **PATCH**: Bug fixes (backward compatible)

## Release Process

1. Update version in `OfficeScrubC2R.psd1`
2. Update `CHANGELOG.md` with release notes
3. Create a tag: `git tag -a v2.19.0 -m "Release v2.19.0"`
4. Push tags: `git push --tags`
5. Create GitHub release
6. Publish to PowerShell Gallery (maintainers only)

## PowerShell Gallery Publishing

For maintainers:

```powershell
# Test the module
Test-ModuleManifest .\OfficeScrubC2R.psd1

# Publish (requires API key)
Publish-Module -Path . -NuGetApiKey $apiKey -Verbose
```

## Questions?

- Open an issue for questions
- Tag with `question` label
- Check existing issues first

## License

By contributing, you agree that your contributions will be licensed under the MIT License.

## Recognition

Contributors will be recognized in:

- README.md (Contributors section)
- Release notes
- Git history

Thank you for contributing! 🎉
