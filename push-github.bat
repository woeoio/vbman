@echo off
chcp 65001 >nul

echo 正在推送到 GitHub...

git push github

echo.
echo ============================================
echo  完成！
echo ============================================
pause
