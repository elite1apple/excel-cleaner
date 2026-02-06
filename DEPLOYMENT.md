# Deployment Guide - Excel Cleaner Web Application

This guide covers deploying your Excel cleaner web application to free hosting platforms.

## Prerequisites

1. **Git & GitHub Account**
   - Install Git: https://git-scm.com/downloads
   - Create GitHub account: https://github.com/signup

2. **Initialize Git Repository** (if not already done)
   ```bash
   cd "c:\Users\mdash\Downloads\Excel-clean"
   git init
   git add .
   git commit -m "Initial commit - Excel Cleaner Web App"
   ```

3. **Create GitHub Repository**
   - Go to https://github.com/new
   - Create a new repository (e.g., "excel-cleaner")
   - Push your code:
     ```bash
     git remote add origin https://github.com/YOUR_USERNAME/excel-cleaner.git
     git branch -M main
     git push -u origin main
     ```

---

## Option 1: Render (Recommended)

**Free Tier:** 750 hours/month, automatic SSL, easy deployment

### Steps:

1. **Sign up at Render**
   - Go to https://render.com/
   - Sign up with GitHub

2. **Create New Web Service**
   - Click "New +" → "Web Service"
   - Connect your GitHub repository
   - Select the `excel-cleaner` repository

3. **Configure Service**
   - **Name:** `excel-cleaner` (or your choice)
   - **Environment:** `Python 3`
   - **Build Command:** `pip install -r requirements.txt`
   - **Start Command:** `gunicorn app:app`
   - **Instance Type:** `Free`

4. **Deploy**
   - Click "Create Web Service"
   - Wait for deployment (5-10 minutes)
   - Your app will be live at: `https://excel-cleaner.onrender.com`

### Updating Your App on Render:
```bash
# Make changes to clean_excel.py or any other files
git add .
git commit -m "Updated cleaning logic"
git push origin main
# Render will automatically redeploy!
```

---

## Option 2: Railway

**Free Tier:** 500 hours/month, $5 initial credit

### Steps:

1. **Sign up at Railway**
   - Go to https://railway.app/
   - Sign up with GitHub

2. **Create New Project**
   - Click "New Project"
   - Select "Deploy from GitHub repo"
   - Choose your `excel-cleaner` repository

3. **Configure**
   - Railway auto-detects Python
   - No additional configuration needed
   - Click "Deploy"

4. **Get Your URL**
   - Go to "Settings" → "Networking"
   - Click "Generate Domain"
   - Your app will be live at: `https://excel-cleaner.up.railway.app`

### Updating Your App on Railway:
Same as Render - just push to GitHub and Railway auto-deploys!

---

## Option 3: PythonAnywhere

**Free Tier:** Limited (good for small usage)

### Steps:

1. **Sign up at PythonAnywhere**
   - Go to https://www.pythonanywhere.com/
   - Create a free account

2. **Upload Files**
   - Click "Files" tab
   - Upload all files from `Excel-clean` directory

3. **Create Web App**
   - Go to "Web" tab
   - Click "Add a new web app"
   - Choose "Flask" framework
   - Set Python version to 3.10

4. **Configure**
   - Set working directory: `/home/YOUR_USERNAME/excel-cleaner`
   - Set WSGI file to point to `app.py`
   - Install requirements:
     ```bash
     pip install -r requirements.txt
     ```

5. **Reload**
   - Click "Reload" button
   - Your app will be live at: `https://YOUR_USERNAME.pythonanywhere.com`

---

## Testing Your Deployment

1. Visit your deployed URL
2. Upload one of the test Excel files from your project
3. Verify the cleaned file downloads correctly
4. Test with different configuration options

---

## Making Updates to Your Application

### For Logic Changes (clean_excel.py):

```bash
# 1. Make your changes to clean_excel.py
# 2. Test locally:
python app.py
# Visit http://localhost:5000

# 3. Commit and push:
git add clean_excel.py
git commit -m "Updated cleaning logic"
git push origin main

# Render and Railway will auto-deploy!
# PythonAnywhere: manually upload new file and click "Reload"
```

### For UI Changes (templates/static):

Same process - edit, commit, push. Your changes will be live in minutes!

---

## Troubleshooting

### "Application Error" or "Service Unavailable"
- Check logs in your hosting platform dashboard
- Common issue: Missing dependencies in `requirements.txt`
- Solution: Verify all imports in `clean_excel.py` are in requirements.txt

### "File Upload Failed"
- Check file size limit (16MB)
- Verify file extension (.xlsx or .xls)
- Check server logs for specific error

### "Download Not Working"
- Check temporary file permissions
- Verify `tempfile` module is working on the platform

### Performance Issues on Free Tier
- Free tiers have limited resources
- Consider upgrading if processing large files frequently
- Render free tier sleeps after inactivity (first request may be slow)

---

## Custom Domain (Optional)

All platforms support custom domains on paid plans:
- **Render:** Settings → Custom Domains
- **Railway:** Settings → Networking → Custom Domain
- **PythonAnywhere:** Paid accounts only

---

## Security Notes

1. **File Validation:** App validates file types and sizes
2. **Temporary Files:** Automatically cleaned up after use
3. **No Database:** No user data is stored
4. **HTTPS:** All platforms provide free SSL certificates

---

## Cost Comparison

| Platform | Free Tier | Best For |
|----------|-----------|----------|
| Render | 750 hrs/month | Most users (recommended) |
| Railway | 500 hrs/month + $5 credit | Power users |
| PythonAnywhere | Limited | Very light usage |

**Recommendation:** Start with Render. It's the easiest and most generous free tier.

---

## Support

For deployment issues:
- **Render:** https://render.com/docs
- **Railway:** https://docs.railway.app/
- **PythonAnywhere:** https://help.pythonanywhere.com/

For application issues:
- Check `app.py` logs
- Verify `clean_excel.py` logic
- Test locally first: `python app.py`
