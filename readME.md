# ExcelArranger
 
## 功能簡介
讀取Excel檔案並透過excel的自定義grouping做分類並另存excel檔案


## 版本
### 下次更新：
1. 分類的group name可顯示（或像目標group ID一樣，列表顯示欄位名稱，讓使用者自行選擇）   當未來表格有變動時，便於自行選擇.
2. 當選擇多個類別時, 因不知處理需多久  
　也許將textbox與button皆設為不可動作較佳.
並且加入loading 畫面
### 0.2.0 ：
 [X] 在選取視窗加入上下左右拉動的bar  
 [X] 能夠一次全選檔案

## How to Pack?
1. use pyinstaller get the spec of file.spec
   1. Pyinstaller -F your_python_code.py  
//-F, --onefile

### OR  

How do I generate a .spec file?  
docker run -v "$(pwd):/src/" cdrx/pyinstaller-linux "pyinstaller your-script.py"  

Create a one-file bundled executable
2. use docker win impement it 
   1. put the .py, .spec file in the src 
   2. and generate the requirement.txt 
   3. docker run -v "$(pwd):/src/" cdrx/pyinstaller-windows:python3
   4. IF any Error check the damn spec
   
