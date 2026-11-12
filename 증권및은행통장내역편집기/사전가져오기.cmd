@echo off
chcp 65001 >nul
rem 화면에서 "사전 파일 저장"으로 내려받은 institutions.js(다운로드 폴더의 가장 최근 파일)를
rem 이 프로젝트의 dictionary\institutions.js 에 반영합니다. 바뀌는 내용을 먼저 보여 주고 확인을 받은 뒤, 백업하고 교체합니다.
cd /d "%~dp0"
set "NODE=node"
where node >nul 2>nul || set "NODE=C:\Program Files\nodejs\node.exe"
"%NODE%" tools\import-dictionary.js %* --pause
