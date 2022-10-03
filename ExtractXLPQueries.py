import os, sys, io, re
import base64, zipfile
import urllib.parse
from shutil import rmtree
from pathlib import Path
import xml.etree.ElementTree as ET
from xml.dom import minidom
from argparse import ArgumentParser

OFFICE_FILE_EXTENSIONS = (
    '.xlsb', '.xlsx', '.xlsm', '.xlam',  # Excelワークブック
)

def get_args():
    parser = ArgumentParser(description='Power Queryのソースコードおよび管理情報の展開')
    parser.add_argument('sources', metavar='MS_OFFICE_FILE', type=str, nargs='+',
                        help='展開するExcelワークブックのファイルまたはディレクトリへのパス.')
    parser.add_argument('--dest', type=str, default='.',
                        help='展開先のディレクトリ [規定値: .].')
    parser.add_argument('--recursive', action='store_true',
                        help='sourcesパラメータにディレクトリが指定されている場合、サブディレクトリを再帰的に検索する.')
    return parser.parse_args()


def get_source_paths(sources, recursive):
    for src in sources:
        p = Path(src)
        # sourceがディレクトリの場合、その下にあるソースファイルを探す
        if p.is_dir(): 
            for file in p.glob("**/*" if recursive else "*"):
                f = Path(file)
                if not f.name.startswith('~$') and f.suffix.lower() in OFFICE_FILE_EXTENSIONS:
                    yield f.absolute()
        # source がファイルの場合、その絶対パスを返す
        else: 
            yield p.absolute()


def file_output(content_string, zpc_name, output_path, dtype):
    
        # ファイル出力パス
        zpc_name = zpc_name.replace('/', os.sep)
        zpc_file = os.path.join(output_path, zpc_name.replace('/', os.sep))
        zpc_path = os.path.dirname(zpc_file)
        if not os.path.exists(zpc_path):
            os.makedirs(zpc_path)

        if dtype == 'text':
            # 改行コードが変換されるのを防ぐ
            with open(zpc_file, "w", newline="", encoding='UTF-8') as f:
                f.write(content_string)
        elif dtype == 'binary':
            with open(zpc_file, "wb") as f:
                f.write(content_string)


def output_queries(queries_lines, queries_output_path):

    # 改行でistに分解する
    queries_lines_list = queries_lines.splitlines()

    #クエリーの先頭パターン
    ptn = '(^shared )(.*)( = )(.*)'

    # クエリーの行中を処理しているならTrue
    in_str = False
    # 行のバッファーの初期化
    str_list = []
    #クエリーの出力ファイルイメージ
    query_files = {}
    
    # ファイルを1行ずつ処理
    for queries_line in queries_lines_list:

        # 先頭行は読み飛ばし
        if queries_line[0:17] == 'section Section1;':
            continue

        if queries_line[0:23] == '[ FunctionQueryBinding ':
            # Functionの定義行
            str_list.append(queries_line + '\n')
            continue

        #今クエリーの行中でなければ
        if not in_str:
            # クエリーの先頭パターンかどうか
            res = re.search(ptn, queries_line)
            if res:
                # 先頭パターンの場合
                query_name = res.group(2)
                if query_name[0] == "#":
                    query_name = query_name[1:].strip('"')
                stn = res.group(4)
                str_list.append(stn + '\n')
                in_str = True
        else:
            # 行を追加    
            str_list.append(queries_line + '\n')
        
        # 最終行だったら
        if queries_line[-1:] == ';':
            # すべての行を辞書に登録する
            query_files[query_name] = str_list
            # 行のバッファーの初期化
            str_list = []
            # クエリーの行の外に設定する
            in_str = False
    
    #ファイル出力
    for query_filename, content_string_list in query_files.items():
        content_string = ''.join(content_string_list)
        file_output(content_string, query_filename + '.pq', queries_output_path, 'text')


def zpc_extract(zpc, package_output_path, queries_output_path):
    
        # メモリー上のzipファイルのByte列をzipファイルとして扱う
        zpc_object = zipfile.ZipFile(io.BytesIO(zpc))
        
        # 各メンバーごとに出力
        for zpc_name in zpc_object.namelist():
            
            # zipファイル内のメンバーの内容
            zpc_content = zpc_object.open(zpc_name, 'r').read()
            zpc_content_string = zpc_content.decode()

            if zpc_name == 'Formulas/Section1.m':
                #クエリーのソースコードごとの出力
                output_queries(zpc_content_string, queries_output_path)
                #ファイル出力
                file_output(zpc_content_string, zpc_name, package_output_path, 'text')
            else:
                #その他パッケージ
                zpc_content_string = pretty_Xml(zpc_content_string) # XMLの場合整形
                #ファイル出力
                file_output(zpc_content_string, zpc_name, package_output_path, 'text')


def pretty_Xml(content_string):
    
    #XMLかどうか
    try:
        ET.fromstring(content_string)
    except ET.ParseError:
        return content_string

    # 整形して返す
    return minidom.parseString(content_string).toprettyxml()


def main(source_file, output_path):
    zip_object  = zipfile.ZipFile(source_file)

    # ワークブックをzipファイルとして格納されているファイルを確認する
    for arcfile in zip_object.namelist():

        # 該当するのは customXml/item数字.xml
        if not re.match(r'^customXml/item[0-9]+.xml$', arcfile):
            continue

        # 該当するファイルを取り出す
        xml_string = zip_object.open(arcfile, 'r').read()
        xml_string = xml_string.decode('utf-16')
        root = ET.fromstring(xml_string)

        # タグがDataMashupのitemが対象
        if root.tag[-10:] == 'DataMashup':
            base64_str = root.text

            # Base64デコード
            bin_stream  = base64.b64decode(base64_str.encode())

            # 解析位置
            zpc_position = 4    # Version情報は不要
            
            ## Package Part

            # 出力ディレクトリ
            package_output_path = os.path.join(output_path, "xml")
            package_output_path = os.path.join(package_output_path, "Package")
            if not os.path.exists(package_output_path):
                os.makedirs(package_output_path)

            queries_output_path = os.path.join(output_path, "QueryCodes")
            if not os.path.exists(queries_output_path):
                os.makedirs(queries_output_path)

            # 4バイトのリトルエンディアンがzipファイルの長さ
            zpc_length = int.from_bytes(bin_stream[zpc_position:zpc_position + 4], byteorder='little')
            zpc_position += 4
            zpc = bin_stream[zpc_position:zpc_position + zpc_length]
            zpc_position += zpc_length
            
            # 展開して出力 (Packageとクエリーの出力先)
            zpc_extract(zpc, package_output_path, queries_output_path)
            
            
            ## Permissions Part

            # 出力ディレクトリ
            Permissions_output_path = os.path.join(output_path, "xml")

            # 4バイトのリトルエンディアンがzipファイルの長さ
            zpc_length = int.from_bytes(bin_stream[zpc_position:zpc_position + 4], byteorder='little')
            zpc_position += 4
            zpc = bin_stream[zpc_position:zpc_position + zpc_length]
            zpc_position += zpc_length
            
            # 展開して出力
            zpc_content_string = zpc.decode()
            zpc_content_string = pretty_Xml(zpc_content_string) # XMLの場合整形
            file_output(zpc_content_string, 'Permissions.xml', Permissions_output_path, 'text')


            # Metadataの長さ (使用しない)
            # Metadata_Length = int.from_bytes(bin_stream[zpc_position:zpc_position + 4], byteorder='little')
            zpc_position += 4

            ## Metadata XML Part

            # 出力ディレクトリ
            MetadataXML_output_path = os.path.join(output_path, "xml")
            MetadataXML_output_path = os.path.join(MetadataXML_output_path, "Metadata")
            if not os.path.exists(MetadataXML_output_path):
                os.makedirs(MetadataXML_output_path)

            # 4バイトのリトルエンディアンがzipファイルの長さ
            zpc_position += 4   # Version情報は不要
            zpc_length = int.from_bytes(bin_stream[zpc_position:zpc_position + 4], byteorder='little')
            zpc_position += 4
            zpc = bin_stream[zpc_position:zpc_position + zpc_length]
            zpc_position += zpc_length
            
            # 展開して出力
            zpc_content_string = zpc.decode()
            
            # <Items><Item><ItemLocation><ItemPath> があったらパーセントエンコーディングを戻す
            root = ET.fromstring(zpc_content_string)
            for itempath in root.findall('./Items/Item/ItemLocation/ItemPath'):
                if itempath.text is not None:
                    itempath_text = itempath.text
                    itempath.text = urllib.parse.unquote(itempath_text, 'utf-8')

            zpc_content_string = pretty_Xml(ET.tostring(root, encoding='utf-8').decode()) # XMLの場合整形
            file_output(zpc_content_string, 'MetadataXML.xml', MetadataXML_output_path, 'text')


            ## Content Part

            # 出力ディレクトリ
            Metadata_Content_output_path = os.path.join(MetadataXML_output_path, "Content")
            if not os.path.exists(Metadata_Content_output_path):
                os.makedirs(Metadata_Content_output_path)

            # 4バイトのリトルエンディアンがzipファイルの長さ
            zpc_length = int.from_bytes(bin_stream[zpc_position:zpc_position + 4], byteorder='little')
            zpc_position += 4
            zpc = bin_stream[zpc_position:zpc_position + zpc_length]
            zpc_position += zpc_length
            
            # 展開して出力
            zpc_extract(zpc, Metadata_Content_output_path, '')
            
            
            ## Permissions Bindings Part

            # 出力ディレクトリ
            Permissions_Bindings_output_path = os.path.join(output_path, "xml")
            Permissions_Bindings_output_path = os.path.join(Permissions_Bindings_output_path, "Permissions_Bindings")
            if not os.path.exists(Permissions_Bindings_output_path):
                os.makedirs(Permissions_Bindings_output_path)

            # 4バイトのリトルエンディアンがzipファイルの長さ
            zpc_length = int.from_bytes(bin_stream[zpc_position:zpc_position + 4], byteorder='little')
            zpc_position += 4
            zpc = bin_stream[zpc_position:zpc_position + zpc_length]
            zpc_position += zpc_length
            
            # バイナリ出力
            file_output(zpc, 'Permissions_Bindings.bin', Permissions_Bindings_output_path, 'binary')

            
            # 処理終了
            break


if __name__ == '__main__':
    args = get_args()
    root = Path(args.dest)
    if not root.exists():
        root.mkdir(parents=True)
    elif not root.is_dir():
        raise FileExistsError

    for source_file in get_source_paths(args.sources, args.recursive):
        src = Path(source_file)
        basename = src.stem
        dest = Path(root.joinpath(basename))
        dest.mkdir(parents=True, exist_ok=True)
        rmtree(dest.absolute(), ignore_errors=True)

        main(src, dest)

