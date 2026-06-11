# Push to GitHub/Sublime

## Step 1: Init git locally
```bash
git init
git add .
git commit -m "Initial commit: Garment Spec OCR FastAPI + React"
```

## Step 2: Create remote repo
Go to https://github.com/new
- Repo name: `spec-ocr` (or whatever)
- Keep empty (don't init with README)

## Step 3: Push
```bash
git remote add origin https://github.com/YOUR_USERNAME/spec-ocr.git
git branch -M main
git push -u origin main
```

## Done
Repo is live at `https://github.com/YOUR_USERNAME/spec-ocr`

---

## For Sublime (or other provider)

Replace GitHub URL with your provider:
- **GitLab:** `https://gitlab.com/YOUR_USERNAME/spec-ocr.git`
- **Gitea:** `https://your-gitea.com/YOUR_USERNAME/spec-ocr.git`
- **Bitbucket:** `https://bitbucket.org/YOUR_USERNAME/spec-ocr.git`
