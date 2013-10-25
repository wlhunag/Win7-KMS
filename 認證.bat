@echo off
echo 本程式將執行Office Professional Plus 2010版本的啟動認證的設定
echo 請確認您執行本程式時是否已經選擇以*系統管理者身分執行*
echo 偵測 Microsoft Office 2010 安裝目錄
set ProgRoot=%ProgramFiles%
echo 設定 KMS 指向中山醫學大學金鑰管理伺服器 - kms.csmu.edu.tw
cscript "%ProgRoot%\Microsoft Office\Office14\ospp.vbs" /sethst:kms.csmu.edu.tw
echo 啟動 Microsoft Office 2010 
cscript "%ProgRoot%\Microsoft Office\Office14\ospp.vbs" /act
echo 啟動程序執行完成
echo 請注意: 上方(約前五行)需有 Product activation successful 出現, 
echo 才表示您的 Office2010 啟動成功!
pause
echo 本程式將執行KMS Win7  Win8 版本的啟動認證的設定
echo 請確認您執行本程式時是否已經選擇以*系統管理者身分執行*

echo 設定 KMS 指向中山醫學大學金鑰管理伺服器 - kms.csmu.edu.tw
slmgr -skms kms.csmu.edu.tw:1688
echo 啟動 Windows
slmgr -ato
echo 啟動程序執行完成
echo 請注意: 需有 啟用成功視窗出現
echo 才表示您的 Windows 啟動成功!
pause
