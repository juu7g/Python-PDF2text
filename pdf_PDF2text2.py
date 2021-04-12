"""
PDFをテキストに変換する
pdfminerを使用
"""

from pdfminer.high_level import extract_pages
from pdfminer.layout import LAParams, LTChar, LTRect, LTTextBox
import datetime
import collections
import os, re, difflib, sys, argparse

class ConvertPDF2csv():
    """
    PDFをtxtに変換する。
    PDFは2段組みの場合も含める
    """

    def __init__(self, argv:list):
        """
        コンストラクタ

        Args:
            argv:   以下
                    入力ファイル名
                    出力ファイル名   拡張子はtxtとする
                    段組みの切れ目   左右の段落の切れ目となる位置    0の場合、罫線情報から計算する
                                    default:1
                    フッター位置     フッターの開始位置     これ以下の文字を変換しない
                                    上記が0の場合、罫線情報から計算する
                    ヘッダー位置     ヘッダーの終了位置     これより上の文字を変換しない
                    開始ページ(1スタート)
                    終了ページ(1スタート)
        """
        self.input_path = r'C:\Users\じゅう\Downloads\70_197.pdf'
        self.output_path = '{}.txt'.format(datetime.datetime.now().strftime("%m%d_%H%M_%S"))
        self.border = 0     # 段組みの切れ目のx座標 0の時ページの真ん中に設定
        self.footer = 60    # フッターのy座標。ページの最下部が0。これより下の位置の文字は抽出しない
        self.header = 1000  # ヘッダーのy座標。 これより上の位置の文字は抽出しない
        self.start_page = 1 # 開始ページ1スタート
        self.last_page = 0  # 終了ページ1スタート

        # コマンドライン引数の解析
        parser = argparse.ArgumentParser()		# インスタンス作成
        parser.add_argument('input_path', type=str, help="入力ファイル名")	# 引数定義
        parser.add_argument('output_path', nargs="?", default=self.output_path, type=str, help="出力ファイル名(default:月日_時分_秒.txt)")	# 引数定義
        parser.add_argument("-b", '--border', type=int, metavar="n", default=1, help="段組みの切れ目  0の場合、用紙幅の半分(default:%(default)s)")	# 引数定義
        parser.add_argument("-f", '--footer', type=int, metavar="n", default=30, help="フッター位置(default:%(default)s)")	# 引数定義
        parser.add_argument("-t", '--top', type=int, metavar="n", default=1000, help="ヘッダー位置(default:%(default)s)")	# 引数定義
        parser.add_argument("-s",'--s_page', type=int, metavar="n", default=1, help="開始ページ(default:%(default)s)")	# 引数定義
        parser.add_argument("-e",'--e_page', type=int, metavar="n", default=0, help="終了ページ(0:最終)(default:%(default)s)")	# 引数定義
        
        if not argv: return
        
        args = parser.parse_args(argv)				# 引数の解析
        print(args)						# 引数の参照
        self.input_path = args.input_path
        self.start_page = args.s_page
        self.last_page = args.e_page
        self.output_path = args.output_path
        self.border = args.border
        self.footer = args.footer
        self.header = args.top
        self.sheet_name = os.path.splitext(os.path.basename(self.output_path))[0]

    def flatten(self, l):
        """
        ツリー状になっているイテレータをフラットに返すイテレータ
        """
        for el in l:
            if isinstance(el, collections.abc.Iterable) and not isinstance(el, (str, bytes)):
                yield from self.flatten(el)
            else:
                yield el

    def flatten_lttext(self, l, _type):
        """
        ツリー状になっているイテレータをフラットに返すイテレータ
        返る要素の型を指定
        pdfminerのextract_pagesで使用するのを想定
        要素の型が引数で指定した型を継承したもののみを返す

        Args:
            l:      pdfminerのextract_pages()の戻り値
            _type:  戻したい値の型
        """
        for el in l:
            if isinstance(el, (_type)):
                yield el
            else:
                if isinstance(el, collections.abc.Iterable) and not isinstance(el, (str, bytes)):
                    yield from self.flatten_lttext(el, _type)
                else:
                    continue

    def write2text(self, f):
        # 1ページ分処理したら書き込む
        f.write(self.text_l)
        f.write(self.text_r)
        self.text_l = self.text_r = ""


    def convert_pdf_to_csv(self):
        """
        PDFファイルをテキストに変換
        PDFは2段に段組みされ、左右で差異を表記した形式の物
        """

        laparams = LAParams()               # パラメータインスタンス
        laparams.boxes_flow = None          # -1.0（水平位置のみが重要）から+1.0（垂直位置のみが重要）default 0.5
        laparams.word_margin = 0.2          # default 0.1
        laparams.char_margin = 2.0          # default 2.0
        laparams.line_margin = 0            # default 0.5

        # 出力ファイルのオープン    ファイルがある時は上書きされる
        with open(self.output_path, "w", encoding="utf-8") as f:
            # 初期化
            self.text_l = ""        # 左側の文字列
            self.text_r = ""        # 右側の文字列
            
            print("Analyzing from {} page to {} page(0:to last)".format(self.start_page, self.last_page))
            
            # 対象ページを読み、テキスト抽出する。（maxpages：0は全ページ）
            for page_layout in extract_pages(self.input_path, maxpages=0, laparams=laparams):    # ファイルにwithしている
                # 抽出するページの選別。extract_pagesの引数では、開始ページだけの指定に対応できないため
                if page_layout.pageid < self.start_page: continue                   # 指定開始ページより前は飛ばす
                if self.last_page and self.last_page < page_layout.pageid: break    # 指定終了ページ以降は中断
                # ページの幅から段組みの境界を計算(用紙幅の半分とする)
                if self.border == 0:
                    self.border = int(page_layout.width / 2)
                if page_layout.pageid == self.start_page:
                    print("Check on page #{}".format(page_layout.pageid))
                    print("Page Info width:{}, heght:{}".format(page_layout.width, page_layout.height))
                    print("Calc result border: {}, footer: {}".format(self.border, self.footer))
                # 要素の出現順の確認(debug)
                # for element in self.flatten_lttext(page_layout, LTTextBox):
                #     print("bbox{} {}".format(element.bbox, element.get_text()[:20]))
                
                # 要素のイテレータをたどり入れ子の要素を1次元に取り出す。戻るイテレータはLTTextBox型のみ
                # 要素の行の上側y1で降順、行の左側x0で昇順にソートする。
                for element in sorted(self.flatten_lttext(page_layout, LTTextBox), key=lambda x: (-x.y1, x.x0)):
                # for element in self.flatten_lttext(page_layout, LTTextBox):
                    if element.y1 < self.footer: continue  # フッター位置の文字は抽出しない
                    if element.y0 > self.header: continue  # ヘッダー位置の文字は抽出しない
                    _text =element.get_text()
                    # debug
                    # print("y1:{}, y0:{}■{}".format(element.y1, element.y0, _text))

                    if element.x1 < self.border:
                        # 文字列全体が左側
                        self.text_l += _text
                    else:
                        if element.x0 >= self.border:
                            # 文字列全体が右側
                            self.text_r += _text
                        else:
                            # 文字列が境界をまたいでいる場合
                            # 右側に既に文章があれば先に出力する
                            if self.text_r:
                                self.write2text(f)
                            self.text_l += _text

                # 1ページ分処理したら書き込む
                self.write2text(f)

if __name__ == "__main__":
    cnv = ConvertPDF2csv(sys.argv[1:])
    cnv.convert_pdf_to_csv()
