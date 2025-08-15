# 🤖 Automation Setup Guide

Complete guide for setting up automated weekly transfers from Kobo to SharePoint.

## 📋 Overview

This guide will help you set up automated weekly execution of the Kobo to SharePoint transfer script. The automation will:

- **🕐 Run every Sunday at 2:00 AM** (configurable)
- **📊 Process all forms automatically** (no user interaction)
- **📝 Log all activities** to `transfer.log`
- **🔄 Handle incremental uploads** (skip existing files)
- **⚡ Resume automatically** if interrupted

## 🎯 Prerequisites

Before setting up automation, ensure:

1. ✅ **Main script works**: `python new.py` runs successfully
2. ✅ **Environment configured**: `.env` file has all required credentials
3. ✅ **Dependencies installed**: All Python packages are available
4. ✅ **Permissions verified**: Both Kobo and SharePoint access work

## 🖥️ Operating System Setup

Choose your operating system below:

---

## 🍎 macOS / Linux Setup

### Option 1: Automated Setup Script (Recommended)

1. **Make script executable**:
   ```bash
   chmod +x setup_weekly_cron.sh
   ```

2. **Run the setup script**:
   ```bash
   ./setup_weekly_cron.sh
   ```

3. **Verify setup**:
   ```bash
   crontab -l
   ```

### Option 2: Manual Cron Setup

1. **Open crontab editor**:
   ```bash
   crontab -e
   ```

2. **Add this line** (replace `/path/to/your/project` with actual path):
   ```bash
   # Run Kobo to SharePoint transfer every Sunday at 2:00 AM
   0 2 * * 0 cd /path/to/your/project && /usr/bin/python3 new_automated.py >> transfer.log 2>&1
   ```

3. **Save and exit** (Ctrl+X, then Y, then Enter)

### Option 3: Custom Schedule

**Daily at 3:00 AM**:
```bash
0 3 * * * cd /path/to/your/project && /usr/bin/python3 new_automated.py >> transfer.log 2>&1
```

**Every Monday at 9:00 AM**:
```bash
0 9 * * 1 cd /path/to/your/project && /usr/bin/python3 new_automated.py >> transfer.log 2>&1
```

**Every 6 hours**:
```bash
0 */6 * * * cd /path/to/your/project && /usr/bin/python3 new_automated.py >> transfer.log 2>&1
```

---

## 🪟 Windows Setup

### Option 1: Automated Setup Script

1. **Run as Administrator**:
   - Right-click `setup_weekly_task.bat`
   - Select "Run as administrator"

2. **Follow the prompts**:
   - The script will create a scheduled task
   - Runs every Sunday at 2:00 AM

### Option 2: Manual Task Scheduler Setup

1. **Open Task Scheduler**:
   - Press `Win + R`
   - Type `taskschd.msc`
   - Press Enter

2. **Create Basic Task**:
   - Click "Create Basic Task"
   - Name: `Kobo SharePoint Transfer`
   - Description: `Weekly automated transfer from Kobo to SharePoint`

3. **Set Trigger**:
   - Select "Weekly"
   - Start: `2:00:00 AM`
   - Recur every: `1 week(s)`
   - Select "Sunday"

4. **Set Action**:
   - Action: "Start a program"
   - Program: `python`
   - Arguments: `new_automated.py`
   - Start in: `C:\path\to\your\project`

5. **Finish**:
   - Check "Open properties dialog"
   - Click Finish

6. **Configure Properties**:
   - General tab: Check "Run whether user is logged on or not"
   - Settings tab: Check "Allow task to be run on demand"
   - Click OK

### Option 3: PowerShell Setup

```powershell
# Create scheduled task via PowerShell
$action = New-ScheduledTaskAction -Execute "python" -Argument "new_automated.py" -WorkingDirectory "C:\path\to\your\project"
$trigger = New-ScheduledTaskTrigger -Weekly -DaysOfWeek Sunday -At 2am
$principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest
$settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries

Register-ScheduledTask -TaskName "Kobo SharePoint Transfer" -Action $action -Trigger $trigger -Principal $principal -Settings $settings
```

---

## 🔧 Configuration Options

### Environment Variables

Ensure your `.env` file is in the project directory:

```env
# Kobo Toolbox API Token
API_TOKEN=your_kobo_api_token_here

# Microsoft 365 / SharePoint Credentials
TENANT_ID=your_tenant_id_here
CLIENT_ID=your_client_id_here
CLIENT_SECRET=your_client_secret_here
SITE_ID=your_sharepoint_site_id_here
```

### Logging Configuration

The automated script (`new_automated.py`) automatically:
- ✅ Logs to `transfer.log` file
- ✅ Shows output in console
- ✅ Includes timestamps
- ✅ Logs errors and warnings

### Custom Logging (Optional)

To customize logging, edit `new_automated.py`:

```python
# Change log level
logging.basicConfig(level=logging.DEBUG)  # More detailed logs

# Change log file name
logging.FileHandler('custom_log.log')

# Add email notifications (requires additional setup)
# Add SMTP configuration for email alerts
```

---

## 📊 Monitoring and Management

### Viewing Logs

**Real-time log monitoring**:
```bash
# macOS/Linux
tail -f transfer.log

# Windows PowerShell
Get-Content transfer.log -Wait
```

**Recent log entries**:
```bash
# Last 50 lines
tail -n 50 transfer.log

# Search for errors
grep "ERROR" transfer.log

# Search for successful transfers
grep "TRANSFER COMPLETE" transfer.log
```

### Manual Execution

**Test the automated script**:
```bash
python new_automated.py
```

**Run with specific forms** (if needed):
```bash
# Use the interactive script for specific forms
python new.py
```

### Managing Scheduled Tasks

**macOS/Linux (Cron)**:
```bash
# View all cron jobs
crontab -l

# Edit cron jobs
crontab -e

# Remove all cron jobs
crontab -r

# Remove specific job (edit and remove the line)
crontab -e
```

**Windows (Task Scheduler)**:
```cmd
# View scheduled tasks
schtasks /query /tn "Kobo SharePoint Transfer"

# Run task manually
schtasks /run /tn "Kobo SharePoint Transfer"

# Delete task
schtasks /delete /tn "Kobo SharePoint Transfer" /f

# Modify task
schtasks /change /tn "Kobo SharePoint Transfer" /tr "python new_automated.py"
```

---

## 🔍 Troubleshooting

### Common Issues

#### 1. Script Not Running
**Symptoms**: No log entries, no files transferred

**Solutions**:
- Check if cron/task scheduler is enabled
- Verify file paths are correct
- Test manual execution: `python new_automated.py`
- Check system time and timezone

#### 2. Permission Errors
**Symptoms**: "Permission denied" in logs

**Solutions**:
- Ensure script has execute permissions: `chmod +x new_automated.py`
- Run setup script as administrator (Windows)
- Check file ownership and permissions

#### 3. Environment Issues
**Symptoms**: "API_TOKEN not found" or authentication errors

**Solutions**:
- Verify `.env` file is in the correct directory
- Check that `.env` file has correct permissions
- Test credentials manually first

#### 4. Network Issues
**Symptoms**: Timeout errors or connection failures

**Solutions**:
- Check internet connectivity
- Verify firewall settings
- Test API endpoints manually
- Consider adding retry logic

### Debug Mode

**Enable detailed logging**:
```python
# Edit new_automated.py
logging.basicConfig(level=logging.DEBUG)
```

**Test with verbose output**:
```bash
python new_automated.py 2>&1 | tee debug.log
```

### Health Check Script

Create a simple health check:

```bash
#!/bin/bash
# health_check.sh

echo "=== Kobo SharePoint Transfer Health Check ==="
echo "Date: $(date)"
echo ""

# Check if .env exists
if [ -f ".env" ]; then
    echo "✅ .env file exists"
else
    echo "❌ .env file missing"
    exit 1
fi

# Check if Python script exists
if [ -f "new_automated.py" ]; then
    echo "✅ new_automated.py exists"
else
    echo "❌ new_automated.py missing"
    exit 1
fi

# Test Python execution
if python3 -c "import requests, pandas, dotenv" 2>/dev/null; then
    echo "✅ Python dependencies available"
else
    echo "❌ Python dependencies missing"
    exit 1
fi

# Test authentication (optional)
echo "Testing authentication..."
python3 -c "
import os
from dotenv import load_dotenv
load_dotenv()
if os.getenv('API_TOKEN') and os.getenv('TENANT_ID'):
    print('✅ Environment variables loaded')
else:
    print('❌ Environment variables missing')
    exit(1)
"

echo ""
echo "✅ Health check completed successfully"
```

---

## 🔄 Maintenance

### Regular Tasks

**Weekly**:
- Review transfer logs for errors
- Check SharePoint folder for new files
- Verify incremental upload is working

**Monthly**:
- Review and rotate API tokens if needed
- Check disk space for log files
- Update Python dependencies if required

**Quarterly**:
- Review and update credentials
- Test manual execution
- Backup configuration files

### Log Rotation

**Automatic log rotation** (Linux/macOS):
```bash
# Add to crontab for monthly log rotation
0 1 1 * * mv transfer.log transfer.log.$(date +%Y%m) && touch transfer.log
```

**Manual log cleanup**:
```bash
# Keep last 10 log files
ls -t transfer.log.* | tail -n +11 | xargs rm
```

---

## 📞 Support

### Getting Help

1. **Check logs first**: `tail -f transfer.log`
2. **Test manually**: `python new_automated.py`
3. **Verify setup**: Run health check script
4. **Review this guide**: Check troubleshooting section

### Emergency Stop

**Stop all automation**:
```bash
# macOS/Linux
crontab -r

# Windows
schtasks /delete /tn "Kobo SharePoint Transfer" /f
```

**Resume automation**:
```bash
# Re-run setup scripts
./setup_weekly_cron.sh  # macOS/Linux
# or
setup_weekly_task.bat   # Windows (as Administrator)
```

---

## 🎉 Success Indicators

When automation is working correctly, you should see:

- ✅ **Regular log entries** every week
- ✅ **New files appearing** in SharePoint
- ✅ **No duplicate uploads** (skipped files in logs)
- ✅ **Successful completion** messages
- ✅ **Consistent folder structure** maintained

---

**🚀 Your automated Kobo to SharePoint transfer system is now ready!**
