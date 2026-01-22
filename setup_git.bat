@echo off
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/re-km/D-3_Sheet_Creater.git
git remote set-url origin https://github.com/re-km/D-3_Sheet_Creater.git
git push -u origin main
