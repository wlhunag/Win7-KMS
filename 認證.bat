@echo off
echo 請以系統管理者身分執行本程式
echo 偵測 Microsoft Office 2010 安裝目錄
set ProgRoot=%ProgramFiles%
echo 設定金鑰管理伺服器，設定為 kms.csmu.edu.tw
cscript "%ProgRoot%\Microsoft Office\Office14\ospp.vbs" /sethst:kms.csmu.edu.tw
echo 啟用 Microsoft Office 2010 
cscript "%ProgRoot%\Microsoft Office\Office14\ospp.vbs" /act
echo 完成啟用
echo 請注意: 上方(約前五行)需有 Product activation successful 出現, 
echo 才表示您的 Office2010 啟動成功!
echo 下一步進行 Win7 或 Win8 的KMS認證
pause
echo 設定中山醫學大學金鑰管理伺服器 - kms.csmu.edu.tw
slmgr -skms kms.csmu.edu.tw:1688
echo 啟用 Windows
slmgr -ato
echo 啟動程序執行完成
echo 請注意: 需有 啟用成功視窗出現
echo 才表示您的 Windows 啟動成功!
pause
