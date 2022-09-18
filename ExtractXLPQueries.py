import os, sys, io
import xml.etree.ElementTree as ET
import base64
import zipfile
import re

def main(source_file):
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

            # 先頭から5～8バイト目の4バイトのリトルエンディアンがzipファイルの長さ
            zpc_length = int.from_bytes(bin_stream[4:8], byteorder='little')
            zpc = bin_stream[8:zpc_length + 8]

            # メモリー上のzipファイルのByte列をzipファイルとして扱う
            zpc_file_object = io.BytesIO(zpc)
            zpc_object = zipfile.ZipFile(zpc_file_object)
            section1 = zpc_object.open('Formulas/Section1.m', 'r').read()
            section1_string = section1.decode()

            # M言語出力
            source_path = os.path.dirname(source_file)
            M_file_path = os.path.join(source_path, "M-Language")
            if not os.path.exists(M_file_path):
                os.makedirs(M_file_path)
            M_file_basename = os.path.splitext(os.path.basename(source_file))[0]
            M_file = os.path.join(M_file_path, M_file_basename + '.m')

            # 改行コードが変換されるのを防ぐ
            with open(M_file, "w", newline="") as f:
                f.write(section1_string)

            # 処理終了
            break


if __name__ == '__main__':

    args = sys.argv

    if 2<= len(args):

        # 解析するワークブックのパス
        source_file = str(args[1])

        main(source_file)

    else:
        print('usage: ExtractQuery.py (ワークブックのパス)')