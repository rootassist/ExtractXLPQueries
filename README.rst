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
- chardet 5.0.0 (導入が必要です)

=====================
Usage
=====================
python ExtractXLPQueries.py [ワークブックのフルパス|ワークブックがあるディレクトリ] [--recursive]  

- パラメーターには対象のワークブックファイルのフルパスか、ワークブックファイルの存在しているディレクトリを指定します
- ディレクトリが指定されているときにパラメーター --recursive が指定されている場合には、ディレクトリを再帰的に検索します
- クエリーを含むすべてのファイルは、ワークブックの存在しているディレクトリの下にワークブックの Basename のディレクトリを作成して展開します
- 各クエリーのソースコードは、上記ディレクトリの下のQueryCodesディレクトリの下に クエリー名+'.pq' というファイル名で展開します  

=====================
License
=====================
このコードはMITライセンスです。自由にご利用ください。
