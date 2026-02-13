# PitchPPT - Frequently Asked Questions

**[English](FAQ.md) | [中文](FAQ.zh-CN.md)**

---

## License

This documentation is part of PitchPPT, licensed under [GNU Affero General Public License v3.0 (AGPLv3)](../LICENSE).

```
Copyright (C) 2024 Jiachen Bao

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU Affero General Public License as published
by the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.
```

---

## General Questions

### Q: What is PitchPPT?

**A:** PitchPPT is a professional pitch deck protection tool that exports your PPT as an image-based PPT. It precisely controls file size while perfectly preserving all annotations, comments, formatting, and animations.

### Q: Is PitchPPT free?

**A:** Yes! PitchPPT is completely free and open-source under the AGPLv3 License. You can use it for personal, educational, and commercial purposes.

### Q: What scenarios is PitchPPT suitable for?

**A:** PitchPPT is ideal for:
- Business pitch decks and investor presentations
- Product launch events
- Academic conference reports
- Any scenario requiring file size limits
- Content protection needs

---

## Installation & Setup

### Q: What are the system requirements?

**A:** 
- Windows 10/11 (64-bit)
- Microsoft PowerPoint 2016 or later
- 4GB RAM minimum (8GB recommended)
- 500MB free disk space

### Q: Do I need to install Microsoft Office?

**A:** Yes, PitchPPT requires Microsoft PowerPoint to be installed. It uses PowerPoint's COM interface for processing.

### Q: Can I use PitchPPT on Mac or Linux?

**A:** Currently, PitchPPT only supports Windows because it relies on the Win32 COM interface. Mac and Linux support may be added in future versions.

### Q: The program won't start. What should I do?

**A:** Try these steps:
1. Ensure PowerPoint is installed and can open normally
2. Run as administrator
3. Check Windows Event Viewer for error details
4. Reinstall Microsoft Office if necessary

---

## Usage Questions

### Q: Which algorithm should I choose?

**A:** 
- **Average Quota (V4)**: Fast, stable, for uniform content
- **Dual-Round Optimization (V5)**: Balanced speed and accuracy
- **Iterative Optimization (V6)**: Best accuracy, for complex PPTs

For competition submissions, we recommend V6 for best results.

### Q: What target size should I set?

**A:** Check your competition requirements:
- Common limits: 20MB, 50MB, 100MB
- Set slightly below limit (e.g., 48MB for 50MB limit)
- Leave margin for unexpected size variations

### Q: Will my annotations be preserved?

**A:** Yes! Unlike traditional compression tools, PitchPPT perfectly preserves:
- All annotations and comments
- Speaker notes
- Revision history
- Formatting and layout

### Q: Can I process multiple files at once?

**A:** Yes! PitchPPT supports batch processing:
1. Select multiple files in the file dialog
2. Or drag and drop multiple files
3. Each file will be processed with the same settings

### Q: What if my PPT has hidden slides?

**A:** You can choose whether to include hidden slides:
- Check "Include hidden slides" to export all slides
- Uncheck to skip hidden slides
- This affects the quota calculation

---

## Technical Questions

### Q: How accurate is the size control?

**A:** PitchPPT typically achieves:
- V4: ±3-5% error
- V5: ±2-3% error
- V6: ±1-2% error

All algorithms ensure the final size does not exceed your target.

### Q: Why does processing take several minutes?

**A:** Processing involves:
1. Opening PowerPoint and loading the file
2. Multiple export and optimization iterations
3. Binary search for optimal parameters
4. Fine-tuning at multiple scales

Complex PPTs with many pages take longer.

### Q: Will image quality be reduced?

**A:** PitchPPT intelligently optimizes image quality:
- Complex pages get higher quality
- Simple pages use lower quality to save space
- Overall visual quality is preserved
- You can adjust DPI settings for better quality

### Q: What file formats are supported?

**A:** 
- **Input**: .pptx, .ppt
- **Output**: .pptx, .pdf, .jpg, .png, .tiff, .bmp

### Q: Can I cancel processing mid-way?

**A:** Yes, click the "Stop" button to cancel. Partial results will be cleaned up automatically.

---

## Troubleshooting

### Q: "PowerPoint initialization failed" error

**A:** Solutions:
1. Ensure PowerPoint is installed
2. Restart your computer
3. Repair Office installation
4. Run PitchPPT as administrator

### Q: "Export failed" error

**A:** Check:
1. Source file is not corrupted
2. Sufficient disk space
3. Output directory is writable
4. No special characters in filename

### Q: Process seems stuck

**A:** 
- Large PPTs (100+ pages) may take 10+ minutes
- Check Task Manager for PowerPoint activity
- If truly stuck, stop and restart

### Q: Output file is still too large

**A:** Try:
1. Use V6 algorithm for better accuracy
2. Set target size lower (e.g., 45MB for 50MB limit)
3. Reduce image quality setting
4. Check for embedded videos or large objects

### Q: Output file quality is poor

**A:** Improve quality:
1. Increase DPI setting (try 200-300)
2. Use V6 algorithm for better allocation
3. Set higher target size if allowed
4. Check original image quality

---

## Best Practices

### Q: How to prepare PPT for best results?

**A:** 
1. Remove unnecessary animations
2. Compress embedded videos separately
3. Use appropriate image resolution
4. Delete unused master slides
5. Save as .pptx (not .ppt)

### Q: What settings for competition submission?

**A:** Recommended:
- Algorithm: V6 (Iterative Optimization)
- Target: 2-3MB below limit
- DPI: 150-200
- Include hidden slides: Based on requirements

### Q: How to verify output quality?

**A:** 
1. Open output in PowerPoint
2. Check all slides in slide sorter view
3. Verify annotations in comments pane
4. Test on different computers

---

## Contributing & Support

### Q: How can I contribute?

**A:** 
- Report bugs on GitHub/Gitee Issues
- Suggest new features
- Submit pull requests
- Share with friends and colleagues

### Q: Where to get help?

**A:** 
- GitHub Issues: https://github.com/baojiachen0214/PitchPPT/issues
- Gitee Issues: https://gitee.com/bao-jiachen/PitchPPT/issues
- Email: thestein@foxmail.com

### Q: How to report bugs effectively?

**A:** Include:
1. PitchPPT version
2. Windows version
3. Office version
4. Error message (screenshot if possible)
5. Steps to reproduce
6. Sample file (if possible)

---

## Future Development

### Q: What features are planned?

**A:** Planned features:
- Mac and Linux support
- Cloud processing
- More output formats
- AI-powered optimization
- Plugin system

### Q: How to stay updated?

**A:** 
- Star the repository on GitHub/Gitee
- Watch for release notifications
- Join our community discussions

---

**Still have questions?** Contact us at thestein@foxmail.com
