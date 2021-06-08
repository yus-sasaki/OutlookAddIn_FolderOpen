# 右クリックメニューでファイルパスからフォルダを開くアドイン　マニュアル

## 目次

* 概要
* 動作仕様
* インストール方法
* 使用方法


## 概要
　このアドインでは、Outlookのメール文章中に存在するファイルパスを右クリックメニューから、そのファイルが直下に格納されている親フォルダを開くことができます。  
　また、ハイパーリンクが途中で切れてしまったフォルダおよびファイルを開きたいという要望を受けて、範囲選択したパスに対し、フォルダおよびファイルを開く機能を追加。
* パスの先頭がIPアドレス：10.71.75.125であった場合、ファイル共有サーバ：helmes-2と単純に置き換えるバージョンを試作しました。私個人の環境ではアクセス権もないため十分なチェックができているか確認がとれず、インストーラーを通常バージョンと分けてあります。

## 動作仕様

### 動作環境
* Outlook2016

### ※注意事項
1．メールに**添付**されているファイルに関して、本アドインを実行することはできません。

2．メールに記載されているファイルパスおよびパス中のフォルダが、そもそも間違っている。または、存在しない場合、フォルダおよびファイルを開くことはできません。

3．メールに記載されているファイルパスおよびパス中のフォルダに対し、アクセス権がない場合、フォルダおよびファイルを開くことはできません。

## インストール方法
1. [Release](https://github.com/yus-sasaki/OutlookAddIn_FolderOpen/releases)の最新バージョンのpublish.zipファイルをダウンロード＆解凍する  

2. 解凍後、setup.exeを実行しインストール
 
### 注意事項
* 追加されたアドインをOutlookのオプションで削除した場合、本アドインをアンインストール後に再度インストールしていただく必要があります。

## 使用方法

### ハイパーリンクが設定されているパスに使用
1．メール文章中のハイパーリンクが設定されたファイルパスおよびフォルダパスに対し、右クリックから右クリックメニューを開きます。

2．メニュー最下部に追加されたボタンをクリックすることで、対応する動作を実行します。

### ハイパーリンクが設定されていないパスに使用
1．メール文章中のファイルパスをドラッグし、範囲選択を行います。

2．範囲選択した領域に対し、右クリックから右クリックメニューを開きます。

3．メニュー最下部に追加されたボタンをクリックすることで、対応する動作を実行します。

### ハイパーリンクが途中で切れてしまったパスに使用
1．メール文章中のハイパーリンクが途中で切れてしまったファイルパスおよびフォルダパスをドラッグし、パスすべてに範囲選択を行います。

2．範囲選択した領域に対し、右クリックから右クリックメニューを開きます。

3．メニュー最下部に追加されたボタンをクリックすることで、対応する動作を実行します。

### 追加されるボタン説明
* **フォルダを開く（ハイパーリンク）**  
  ハイパーリンクが設定されたファイルパスのファイルが直下に格納されているフォルダを開きます。   
  また、ハイパーリンクがフォルダに設定されている場合、ハイパーリンクのフォルダを開きます。  
  （ハイパーリンクが途中で切れている場合、設定が有効な部分までのフォルダを開きます。）

* **フォルダを開く（範囲選択）**    
  ドラッグし、範囲選択を行ったファイルパスのファイルが直下に格納されているフォルダを開きます。  
  また、範囲選択を行ったパスがフォルダの場合、範囲選択したフォルダパスのフォルダを開きます。  
  ハイパーリンクが設定されていないフォルダ、ハイパーリンクが途中で切れてしまったフォルダにご使用ください。

* **ファイルを開く（範囲選択）**  
  ドラッグし、範囲選択を行ったファイルパスのファイルを開きます。  
  ハイパーリンクが設定されていないファイル、ハイパーリンクが途中で切れてしまったファイルにご使用ください。

