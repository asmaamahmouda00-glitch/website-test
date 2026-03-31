# 🚀 Quick Start Guide

## Installation (One-Time Setup)

```bash
# Install dependencies
pip install -r requirements.txt
```

## Running the Automation

### Option 1: Test with Your Specific Account Data ⭐ (RECOMMENDED)

```bash
python test_with_custom_data.py
```

This uses your provided account:
- Email: `ai.dexter85@worldposta.com`
- Name: `AI dexter85`
- Company: `AI Company dexter85`
- Phone: `1095666032`

### Option 2: Auto-Generated Random Test Data

```bash
python worldposta_automation.py
```

Generates random test accounts automatically.

### Option 3: Batch Processing Multiple Accounts

```bash
python batch_runner.py
```

1. First run creates `accounts_to_register.csv`
2. Edit the CSV with your accounts
3. Run again to process all accounts

## What Happens During Automation?

1. ✅ **Registration** - Fills and submits registration form
2. ✅ **Email Login** - Logs into WorldPosta webmail
3. ✅ **Find Email** - Searches for "Welcome To WorldPosta Business Email"
4. ✅ **Verify** - Clicks verification link
5. ✅ **Website Login** - Logs into admin panel
6. ✅ **Actions** - Clicks "View Posta" and "View CloudEdge"
7. ✅ **Screenshot** - Captures final result
8. ✅ **Logs** - Saves to CSV and JSON

## Output Files

📸 **Screenshots:** `C:\Users\olaaa\Desktop\Projects\Registeration\SS\`
📊 **CSV Log:** `registration_results.csv`
📋 **JSON Log:** `registration_results.json`

## Troubleshooting

### ❌ Module not found
```bash
pip install -r requirements.txt
```

### ❌ Email not found (timeout)
- Wait longer (script waits 300 seconds by default)
- Check email was sent correctly
- Verify email subject line is correct

### ❌ Element not found
- Website HTML may have changed
- Check selectors in the script
- Contact support for updated selectors

### ❌ Screenshot path error (Windows)
Make sure path uses raw string:
```python
SCREENSHOT_DIR = r"C:\Users\olaaa\Desktop\Projects\Registeration\SS"
```

## Configuration

Edit `worldposta_automation.py` to customize:

```python
# Timeouts
EMAIL_WAIT_TIMEOUT = 300  # Seconds to wait for verification email
DEFAULT_TIMEOUT = 30      # WebDriverWait timeout

# Screenshot directory
SCREENSHOT_DIR = r"C:\Your\Custom\Path"

# Output files
CSV_FILE = "registration_results.csv"
JSON_FILE = "registration_results.json"
```

## Tips

✅ **First Run:** Use visible browser mode (default) to watch the automation
✅ **Debugging:** Check screenshots if any step fails
✅ **Scaling:** Use batch_runner.py for multiple accounts
✅ **Delays:** Script uses human-like random delays (1-8 seconds)
✅ **Detection:** Uses undetected-chromedriver to avoid bot detection

## Need Help?

1. Check console output for error messages
2. Review error screenshots in SS folder
3. Check CSV/JSON logs for detailed status
4. Verify Chrome browser is installed and updated
5. Ensure all selectors match current website HTML

---

**Ready to automate! 🎯**
