@echo off
echo ============================
echo Navigating to project folder
echo ============================
cd /d D:\New_Jira_Data_Extractor

echo ============================
echo Switching to 'master' branch
echo ============================
git checkout master

echo ============================
echo Pulling latest changes
echo ============================
git pull origin main

echo ============================
echo Staging changes
echo ============================
git add .

echo ============================
echo Committing changes if any
echo ============================
git diff --cached --quiet
IF %ERRORLEVEL% EQU 1 (
    git commit -m "Auto commit from Jenkins"
    echo Commit done.
) ELSE (
    echo No changes to commit.
)

echo ============================
echo Pushing to GitHub
echo ============================
git push origin master

echo ============================
echo Git push completed.
echo ============================
pause
