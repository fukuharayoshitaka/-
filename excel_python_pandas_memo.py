import pandas as pd
from glob import glob
filepaths = glob('source/*.xlsx')
filepaths
filepath = filepaths[0]
filepath

def extract(filepath):
    _df = pd.read_excel(filepath)
    columns = _df.iloc[10, [1, 2, 4, 10, 11, 14]]
    df = _df.iloc[11:23, [1, 2, 4, 10, 11, 14]]
    df['企業名'] = _df.iloc[2, 0]
    df['請求書No'] = _df.iloc[2, 12]
    df['発行日'] = _df.iloc[3, 12]
    df['発行者'] = _df.iloc[4, 12]
    df['うんちちコード'] = _df.iloc[4, 13]
    return df
    
df = extract(filepath)
df.head()
sample_1 = pd.DataFrame([[1, 2, 3], [4, 5, 6]])
sample_2 = pd.DataFrame([[-1, 0, 2], [8, 4, 9]])
sample_2
pd.concat([sample_1, sample_2])
df = pd.DataFrame()

df
for filepath in filepaths:
    _df = extract(filepath)
    df= pd.concat([df, _df])
df.dropna()
df.tail()
df = df.iloc[:, [6, 7, 8, 9, 10, 11, 0, 1, 2, 3, 4, 5]]
df.to_excel('output/all_data.xlsx', index=False)

### pandas.DataFrame, pandas.Series
列名のリストを指定すると、選択した複数列をpandas.DataFrameとして取得できる。列の順番は指定したリストの順番になる。

df[['point', 'age']]

### glob  
Pythonのglobモジュールを使うと、ワイルドカード*などの特殊文字を使って条件を満たすファイル名・ディレクトリ（フォルダ）名などの  
パスの一覧をリストやイテレータで取得できる。

### pd.read_ecxel(filepath)  ###pd.read_csv
pandasでExcelファイル（拡張子:.xlsx, .xls）をpandas.DataFrameとして読み込む
 - 一つのシートを読み込み  
 - 複数のシートを読み込み  
 - すべてのシートを読み込み  
#### pd.read_csv('.csv', index_col=0)
headerとindex（見出し列）がある以下のようなcsvファイルを読み込む。
#   Unnamed: 0   a   b   c   d
# 0        ONE  11  12  13  14
# 1        TWO  21  22  23  24
# 2      THREE  31  32  33  34

### columns = _df.iloc[]
columnsは列名（列ラベル）、indexは行名（行ラベル）。valuesは実際のデータの値、
pandas.DataFrameの任意の位置のデータを取り出したり変更（代入）したりするには、at, iat, loc, ilocを使う
  - 位置の指定方法
      at, loc : 行名（行ラベル）、列名（列ラベル）
      iat, iloc : 行番号、列番号
  - 選択し取得・変更できるデータ
      at, iat : 単独の要素の値
      loc, iloc : 単独および複数の要素の値
###pd.header()
行数の多いpandas.DataFrame, pandas.Seriesのデータを確認するときに、先頭（最初）と末尾（最後）の行を返すメソッドhead()とtail()が便利。
引数に数字を入れて数を指定　初期値５
### pd.concat()  
データの単純な結合
 pd.concat([df1, df2, df1])何個でも連結可能
  axisで連結方法が縦横と変わるが注意
### df.dropna()
pandas.DataFrame, Seriesの欠損値NaNを削除（除外）するにはdropna()メソッドを使う。
