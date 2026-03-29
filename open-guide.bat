@echo off
echo 캘리버 전자책 가이드를 엽니다...
start http://localhost:4000
cd /d D:\Sites\calibre-guide\docs
npx docsify-cli serve . --port 4000
