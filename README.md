# LinstringsGeneratorForTrafficCensus
Programs that apply traffics data between railroad station to the vector data of railroad network.  
  
# 概要
国土交通省総合政策局実施の「大都市交通センサス」の調査結果集計表「【鉄道調査】報告書資料編（EXCEL形式）」3.駅別発着・駅間通過人員表のデータを
鉄道路線図のベクターデータの属性として付与するための一連のプログラム。
ベクターデータは国土地理院の「国土数値情報」の「鉄道データ」から取得。  
  

# 実行環境
 - 言語  
  Python3.7.0
 - アプリケーション  
  QGIS 3.10.14  PythonはQGISをインストールするとついているpython3.exeを利用。  
   
# プログラム
  - railroadtraffic.py  
   クラス「RailRoadsLinesWithTraffic」を定義している
  
  - qgis_init.py  
   railroadtraffic.py 初期化時に各種環境変数を定義する関数「setenvirons」を定義
    
  - createtrafficlayers.py  
   railroadtraffic.pyのクラスをインポートし、国土数値情報「鉄道データ」を加工し、作業用のShapefileを作成するプログラム。
    
  - createworklayers.py  
   railroadtraffic.pyのクラスをインポートし、作業用のshapefileに大都市交通センサスのデータを属性に付与したshapefileを出力するプログラム。
   路線探索エンジンは独自開発。  
 
# データ
   フォルダー work 内に保存。
   - 12_駅別発着・駅間通過人員表_首都圏.xlsx  
  「大都市交通センサス」の調査結果集計表「【鉄道調査】報告書資料編（EXCEL形式）」3.駅別発着・駅間通過人員表（首都圏）のデータを
   加工したもの。ワークシート「鉄道定期券合計+普通券」のうち、shapefileの属性として記入されるのはN列～S列のデータ（合計）。
   T列～X列に「国土数値情報」の「鉄道データ」における駅名・鉄道会社名および路線名を記述（大都市交通センサスのそれらとは表記が異なるため）。
    
   - RailroadSection_2.shp  
   プログラム「createtrafficlayers.py」が出力したshapefile（ラインストリング）を加工したもの。
   加工内容としては、「国土数値情報」の「鉄道データ」において路線が途中で途切れているもの等の書き足し、
   実際の電車の運転系統を考慮した路線ラインストリングの追加等。
   
   - RailroadStation_2.shp  
  プログラム「createtrafficlayers.py」が出力したshapefile（ポイント）を加工したもの。
   ポイントは駅のラインストリングの中間地点に配置している。
   加工内容としては、実際の電車の運転路線等を考えた経路指定のためのダミーポイントの追加など。  
   
# 元データの出どころ  
  - 国土交通省総合政策局実施の「大都市交通センサス」の調査結果集計表「【鉄道調査】報告書資料編（EXCEL形式）」3.駅別発着・駅間通過人員表（首都圏）のデータ  
    https://www.mlit.go.jp/sogoseisaku/transport/sosei_transport_tk_000035.html
  - 国土数値情報の「鉄道データ」  
    https://nlftp.mlit.go.jp/ksj/gml/datalist/KsjTmplt-N02-v2_3.html
    
# 出来上がりデータのビュー
  プログラム「createworklayers.py」が出力したラインストリングShapefileを編集しKMLを作成のうえ、Googleマップにインポート。以下のサイトを参照のこと。  
  https://sites.google.com/view/jizai-kobo/%E3%83%9B%E3%83%BC%E3%83%A0/%E5%A4%A7%E9%83%BD%E5%B8%82%E4%BA%A4%E9%80%9A%E3%82%BB%E3%83%B3%E3%82%B5%E3%82%B9
