# GitHub Release Creator Script
# Usage: .\create_github_release.ps1 -Token "your_token"

param(
    [Parameter(Mandatory=$true)]
    [string]$Token
)

$owner = "baojiachen0214"
$repo = "PitchPPT"
$tag = "v1.6.0"
$releaseName = "PitchPPT v1.6.0 - Initial Release"
$releaseBody = @"
## 🎉 Initial Release

We are excited to announce the first official release of PitchPPT!

## ✨ Features

### Core Functionality
- **Three Intelligent Optimization Algorithms**
  - Average Quota Algorithm: Equal quota per page, fast & stable
  - Dual-Round Optimization Algorithm: Test then adjust, balanced approach
  - Iterative Optimization Algorithm: Complexity-based allocation, highest accuracy (<2% error)

- **Two Processing Modes**
  - Standard Mode: Quick export with unified quality settings
  - Smart Mode: Precise file size control with intelligent algorithms

- **Batch Processing**: Process multiple files simultaneously with progress tracking

### Image Options
- **Multiple Formats**: PNG, JPEG, TIFF, WebP, BMP
- **DPI Presets**: 72 (Screen) to 600 (Ultra HD) - up to 16K resolution
- **Smart Resolution Range**: 480px to 4000px height (auto-optimized)

### Content Protection
- Export as image-based PPT (each slide becomes background image)
- Complete structure preservation:
  - Annotations and comments
  - Slide transitions and animations
  - Speaker notes
  - Hyperlinks

## 📦 Installation

### System Requirements
- Windows 10/11 (64-bit)
- Microsoft PowerPoint 2016 or later

### Download
- **File**: PitchPPT.exe (≈47 MB)

## 🚀 Quick Start

1. Run PitchPPT.exe
2. Select Standard Mode or Smart Mode
3. Add your PPT file(s)
4. Configure settings and start conversion

## 📄 License

This project is licensed under GNU AGPLv3.

## 🔗 Links

- GitHub: https://github.com/baojiachen0214/PitchPPT
- Gitee: https://gitee.com/bao-jiachen/PitchPPT

---

**Author**: Jiachen Bao
**Contact**: thestein@foxmail.com
"@

# Create Release
$headers = @{
    "Authorization" = "token $Token"
    "Accept" = "application/vnd.github.v3+json"
}

$body = @{
    tag_name = $tag
    name = $releaseName
    body = $releaseBody
    draft = $false
    prerelease = $false
} | ConvertTo-Json

try {
    Write-Host "Creating release..."
    $response = Invoke-RestMethod -Uri "https://api.github.com/repos/$owner/$repo/releases" -Method Post -Headers $headers -Body $body -ContentType "application/json"
    Write-Host "Release created successfully!"
    Write-Host "Upload URL: $($response.upload_url)"
    
    # Upload asset
    $uploadUrl = $response.upload_url -replace "{\\?name,label}", "?name=PitchPPT.exe"
    $filePath = "dist\PitchPPT.exe"
    
    Write-Host "Uploading PitchPPT.exe..."
    $fileBytes = [System.IO.File]::ReadAllBytes($filePath)
    $uploadHeaders = @{
        "Authorization" = "token $Token"
        "Accept" = "application/vnd.github.v3+json"
        "Content-Type" = "application/octet-stream"
    }
    
    $uploadResponse = Invoke-RestMethod -Uri $uploadUrl -Method Post -Headers $uploadHeaders -Body $fileBytes
    Write-Host "File uploaded successfully!"
    Write-Host "Download URL: $($uploadResponse.browser_download_url)"
}
catch {
    Write-Host "Error: $_" -ForegroundColor Red
}
