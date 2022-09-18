=====================
Name
=====================
ワークブックからPower Queryの情報を取り出す

=====================
Overview
=====================
ExcelのワークブックからPower Queryのソースコード(M言語)やその他の管理情報を取り出します

=====================
Description
=====================
Microsoftが発行している以下の技術文書に基づいて、ExcelのワークブックからPower Queryのソースコード(M言語)やその他の管理情報を取り出します 

- https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-qdeff/27b1dd1e-7de8-45d9-9c84-dfcc7a802e37

pythonは以下のバージョンで動作確認しています(動作確認はWindowsでしか行っていません)

- python 3.10.4  

=====================
Usage
=====================
python ExtractXLPQueries.py (ワークブックのフルパス)  

- ワークブックの存在しているディレクトリの下にワークブック名のディレクトリの下に展開されます
- Power Queryのソースコードは以下のファイルにすべてのクエリーを結合した形で展開されます  
  Package/Formulas/Section1.m

=====================
ライセンス
=====================
このコードはMITライセンスです。自由にご利用ください。
