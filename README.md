# Python-PDF2text

## 概要 Description
PDFファイルを読んで文字をテキストファイルに出力します。  

## 特徴 Features
ページのヘッダーやフッターを抽出の対象から除けます。  
ページを指定して抽出できます。 
2段組みの文書でも抽出できます。  

## 依存関係 Requirement

- Python 3.8.5
- pdfminer.six 20201018

## 使い方 Usage

```dosbatch
usage: pdf_PDF2text.exe [-h] [-b n] [-f n] [-t n] [-s n] [-e n]
                        input_path [output_path]

positional arguments:
  input_path        入力ファイル名
  output_path       出力ファイル名(default:月日_時分_秒.txt)

optional arguments:
  -h, --help        show this help message and exit
  -b n, --border n  段組みの切れ目 0の場合、用紙幅の半分(default:1)
  -f n, --footer n  フッター位置(default:30)
  -t n, --top n     ヘッダー位置(default:1000)
  -s n, --s_page n  開始ページ(default:1)
  -e n, --e_page n  終了ページ(0:最終)(default:0)
```

## インストール方法 Installation

- pip install pdfminer.six

## 作者 Authors
juu7g

## ライセンス License
This software is released under the MIT License, see LICENSE.txt.
（このソフトウェアは、MITライセンスのもとで公開されています。LICENSE.txtを堪忍してください。） 

## References
参考にした情報源（サイト・論文）などの情報、リンク
