# PitchPPT Troubleshooting Guide

**[English](TROUBLESHOOTING.md) | [中文](TROUBLESHOOTING.zh-CN.md)**

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

This guide helps you resolve common issues when using PitchPPT.

---

## Installation Issues

### Issue: "ModuleNotFoundError: No module named 'src'"

**Cause:** Running the file from wrong directory.

**Solution:**
```bash
# Correct way
cd d:\Pythoncode\PitchPPT
python src\main.py

# Incorrect way (don't do this)
cd d:\Pythoncode\PitchPPT\src
python main.py
```

### Issue: "ImportError: No module named 'PyQt5'"

**Cause:** Dependencies not installed.

**Solution:**
```bash
pip install -r requirements.txt
```

### Issue: "ImportError: No module named 'win32com'"

**Cause:** pywin32 not installed.

**Solution:**
```bash
pip install pywin32
```

---

## PowerPoint Issues

### Issue: "PowerPoint initialization failed"

**Symptoms:**
- Error message about PowerPoint COM interface
- Program crashes on start

**Solutions:**

1. **Verify PowerPoint Installation**
   ```
   - Open PowerPoint manually
   - Check if it opens without errors
   - Try creating a new presentation
   ```

2. **Repair Office Installation**
   ```
   Windows Settings → Apps → Microsoft Office → Modify → Repair
   ```

3. **Run as Administrator**
   ```
   Right-click PitchPPT.exe → Run as administrator
   ```

4. **Check PowerPoint Version**
   - Minimum required: PowerPoint 2016
   - Recommended: PowerPoint 2019/365

### Issue: "PowerPoint is busy"

**Cause:** Another PowerPoint instance is running.

**Solution:**
1. Close all PowerPoint windows
2. Check Task Manager for POWERPNT.EXE
3. End any lingering PowerPoint processes
4. Restart PitchPPT

---

## Conversion Issues

### Issue: "Export failed"

**Common Causes & Solutions:**

1. **Corrupted Source File**
   ```
   - Open PPT in PowerPoint
   - Save as new file
   - Try converting the new file
   ```

2. **Insufficient Disk Space**
   ```
   - Check available space (need 2x target size)
   - Clean up temporary files
   - Change output directory
   ```

3. **Permission Issues**
   ```
   - Run as administrator
   - Change output to user directory (Desktop, Documents)
   - Check folder permissions
   ```

4. **Special Characters in Filename**
   ```
   - Remove special characters: / \ : * ? " < > |
   - Use simple ASCII filenames
   ```

### Issue: "Process stuck at X%"

**Diagnosis:**
1. Check if PowerPoint is still active in Task Manager
2. Look at the log file for details
3. Check if file is unusually large

**Solutions:**

1. **Wait Longer**
   - Large files (100+ pages) can take 10+ minutes
   - Complex animations slow down processing

2. **Reduce Complexity**
   - Remove unnecessary animations
   - Compress embedded videos
   - Simplify complex graphics

3. **Restart and Retry**
   ```
   - Stop the process
   - Restart PitchPPT
   - Try with different settings
   ```

### Issue: "Output file size exceeds target"

**Causes:**
1. Algorithm couldn't compress enough
2. Target size too close to limit
3. File contains non-compressible elements

**Solutions:**

1. **Use V6 Algorithm**
   - Most accurate size control
   - Better quota allocation

2. **Lower Target Size**
   ```
   If limit is 50MB, set target to 45-47MB
   ```

3. **Check for Large Elements**
   ```
   - Embedded videos (convert to linked)
   - High-res images (compress separately)
   - Embedded fonts (subset fonts)
   ```

4. **Reduce Quality Settings**
   ```
   - Lower DPI (try 96-150)
   - Reduce image quality slider
   ```

---

## Quality Issues

### Issue: "Images look blurry"

**Solutions:**

1. **Increase DPI**
   ```
   Settings → Image Quality → DPI: 200-300
   ```

2. **Use V6 Algorithm**
   - Better allocation for complex pages
   - Preserves detail where needed

3. **Increase Target Size**
   ```
   If competition allows, use higher limit
   ```

4. **Check Original Quality**
   ```
   - Original images might be low quality
   - Replace with higher resolution versions
   ```

### Issue: "Text looks pixelated"

**Cause:** Text rendered as image at low resolution.

**Solutions:**

1. **Keep Text as Text**
   ```
   - Don't convert text to images
   - Use PowerPoint's native text boxes
   ```

2. **Increase Export Resolution**
   ```
   Higher DPI preserves text sharpness
   ```

3. **Use Vector Graphics**
   ```
   - Use SVG logos instead of PNG
   - Keep charts as editable objects
   ```

---

## UI Issues

### Issue: "Window doesn't display correctly"

**Solutions:**

1. **Check Display Scaling**
   ```
   Windows Settings → Display → Scale and Layout
   - Try 100% or 125%
   ```

2. **Update Graphics Drivers**
   ```
   Device Manager → Display adapters → Update driver
   ```

3. **Run in Compatibility Mode**
   ```
   Right-click exe → Properties → Compatibility
   ```

### Issue: "Buttons not responding"

**Causes:**
- Background process blocking UI
- Thread deadlock

**Solutions:**

1. **Wait for Current Operation**
   - Some operations block UI temporarily
   - Check progress bar

2. **Restart Application**
   ```
   Close and reopen PitchPPT
   ```

3. **Check Logs**
   ```
   logs/pitchppt_*.log for errors
   ```

---

## Performance Issues

### Issue: "Processing is very slow"

**Optimization Tips:**

1. **Close Other Applications**
   - Free up RAM and CPU
   - Especially other Office apps

2. **Use SSD for Temp Files**
   ```
   Set temp directory to SSD drive
   ```

3. **Reduce PPT Complexity**
   ```
   - Remove unused master slides
   - Compress images before processing
   - Delete hidden slides if not needed
   ```

4. **Choose Faster Algorithm**
   ```
   V4 is fastest, V6 is slowest
   ```

### Issue: "High memory usage"

**Normal Behavior:**
- Peak usage: 500MB-1GB for large files
- Temporary during processing

**If Excessive:**

1. **Process Smaller Batches**
   ```
   Don't process 100+ files at once
   ```

2. **Restart Between Sessions**
   ```
   Close and reopen to free memory
   ```

3. **Check for Memory Leaks**
   ```
   Monitor memory in Task Manager
   Report if continuously increasing
   ```

---

## Error Codes

### Error: "Optimization failed: Base volume A exceeds target"

**Meaning:** Your PPT without images is already larger than target.

**Solutions:**
1. Increase target size
2. Remove non-image content
3. Compress embedded objects

### Error: "Export failed: Invalid file format"

**Meaning:** Source file is corrupted or not a valid PPT.

**Solutions:**
1. Open and re-save in PowerPoint
2. Export to new PPTX file
3. Check file extension

### Error: "Permission denied"

**Meaning:** Cannot read source or write to destination.

**Solutions:**
1. Run as administrator
2. Change output directory
3. Check file is not open elsewhere

---

## Getting Help

If issues persist:

1. **Check Logs**
   ```
   logs/pitchppt_*.log
   ```

2. **Enable Debug Mode**
   ```
   Set logging level to DEBUG in config
   ```

3. **Contact Support**
   - GitHub Issues: https://github.com/baojiachen0214/PitchPPT/issues
   - Email: thestein@foxmail.com

4. **Provide Information**
   - PitchPPT version
   - Windows version
   - Office version
   - Error messages
   - Log files

---

**Remember:** Most issues can be resolved by restarting PowerPoint and PitchPPT!
