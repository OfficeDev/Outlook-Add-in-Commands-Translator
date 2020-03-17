---
page_type: sample
products:
- office-outlook
- office-365
languages:
- javascript
extensions:
  contentType: samples
  technologies:
  - Add-ins
  createdDate: 8/21/2015 10:55:59 AM
---
# Outlook-Add-in-Commands-Translator

Translator アドインは、Outlook アドインのコマンド モデルを使用して、新しいメッセージ フォーム内のリボンにボタンを追加します。このボタンは、メッセージ本文で選択したテキストを翻訳 Web サービスに送信して、英語からロシア語に翻訳します。

## 前提条件

このサンプルを実行するには、次のものが必要です。

- サンプル ファイルをホストする Web サーバー。サーバーは、SSL で保護された要求 (https) を受け入れることが可能で、有効な SSL 証明書を所有している必要があります。
- Office 365 メール アカウント **または** Outlook.com メール アカウント。[Office 365 Developer プログラムに参加すると、Office 365 の 1 年間無料のサブスクリプションを取得](https://aka.ms/devprogramsignup)できます。
- [Yandex Translate API](https://translate.yandex.com/developers) の API キー。
- [Office 2016 プレビュー](https://products.office.com/en-us/office-2016-preview)の一部である Outlook 2016。

## サンプルの構成とインストール

1. レポジトリをダウンロードまたはフォークします。
1. テキスト エディターで `TranslateHelper.js` を開き、`YOUR API KEY HERE` の値を Yandex Translate API の API キーで置き換えます。変更内容を保存します。
1. `AppCompose`、`Content`、および `Images` の各ディレクトリを、Web サーバーのディレクトリにアップロードします。
1. `TranslateAppManifest.xml` をテキスト エディターで開きます。`YOUR_WEB_SERVER` のすべてのインスタンスを、前の手順でファイルをアップロードした先のディレクトリの HTTPS URL で置き換えます。変更内容を保存します。
1. ブラウザーを使用して、https://outlook.office365.com (Office 365 用) または https://www.outlook.com (Outlook.com 用) のいずれかでメール アカウントにログオンします。右上隅にある歯車アイコンをクリックし、 [**アドインの管理**] をクリックします。
    
  ![https://www.outlook.com の [アドインの管理] メニュー項目](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-manage-addins.PNG)
    
1. アドインの一覧で [**+**] アイコンをクリックし、[**ファイルから追加**] を選択します。

  ![アドインの一覧の [ファイルから追加] メニュー項目](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/addin-list.PNG)

1. [**参照**] をクリックし、開発用コンピューター上の [`TranslateAppManifest.xml`] ファイルを参照します。[**次へ**] をクリックします。

  ![[ファイル からアドインを追加] ダイアログ](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/browse-manifest.PNG)

1. アドインが Office Store からのものではなく、Microsoft により確認されていないという警告が確認画面に表示されます。[**インストール**] をクリックします。
1. 次の成功メッセージが表示されます:**Outlook 用のアドインを追加しました**。[OK] をクリックします。

## サンプルの実行 ##

1. Outlook 2016 を開き、アドインをインストールしたメール アカウントに接続します。
1. 新しいメール メッセージを作成します。アドインにより [**Translate (翻訳)**] ボタンがコマンド リボンに配置されたことを確認してください。

  ![Outlook の新しいメール フォームの [Translate (翻訳)] ボタン](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/new-mail.PNG)

1. 本文に英語のテキストを入力します。テキストを選択し、[**Translate (翻訳)**] をクリックします。

  ![本文で英語のテキストが選択された新しいメール フォーム](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-selected.PNG)

1. しばらくすると、選択したテキストがロシア語の翻訳に置き換わり、"**Translated successfully (正常に翻訳されました)**" というメッセージが情報バーに表示されます。

  ![アドインによってロシア語に翻訳されたテキスト](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-translated.PNG)

## サンプルの主な構成要素

- [```TranslateAppManifest.xml```](TranslateAppManifest.xml):Translator アドインのマニフェスト ファイル。
- [```AppCompose/FunctionFile/Home.html```](AppCompose/FunctionFile/Home.html):アドイン コマンドをサポートするクライアント用の、`Functions.js` を読み込むための空の HTML ファイル。
- [```AppCompose/FunctionFile/Translator.js```](AppCompose/FunctionFile/Translator.js):アドイン コマンド ボタンがクリックされたときに呼び出されるコード。
- [```AppCompose/Home/Home.html```](AppCompose/Home/Home.html):アドイン コマンドをサポートしていないクライアントにより読み込まれて表示される HTML ファイル。
- [```AppCompose/Home/Home.js```](AppCompose/Home/Home.js):アドイン コマンドをサポートしていないクライアントにより呼び出されるコード。
- [```AppCompose/TranslateHelper.js```](AppCompose/TranslateHelper.js):`Translator` と `Home.js`の両方で使用される共通コード。

## 動作の仕組み

このサンプルの重要な部分は、マニフェスト ファイルの構造です。Office アドインのすべてのマニフェストと同様、このマニフェストではバージョン 1.1 のスキーマが使用されています。ただし、このマニフェストには `VersionOverrides` という新しいセクションがあります。このセクションには、アドイン コマンドをサポートしているクライアント (現在は Outlook 2016 のみ) がアドインをリボン ボタンから呼び出すのに必要なすべての情報が含まれています。この情報を完全に別のセクションに含めることにより、元の新規作成モードのアドインモデルの下の作業ウィンドウへのアドインの読み込みを可能にする、元のマークアップをマニフェストに含めることもできます。アドインを Outlook 2013 または Outlook on the web に読み込むと、この動作を確認することができます。

### Outlook on the web に読み込まれた Translator アドイン ###

![Outlook on the web に読み込まれた Translator アドイン](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-on-web.PNG)

`VersionOverrides` 要素内には、`Resources` と `Hosts` という 2 つの子要素があります。`Resources` 要素には、アドイン用のアイコン、文字列、および読み込む HTML ファイルに関する情報が含まれています。`Hosts` セクションは、アドインが読み込まれる方法とタイミングを指定します。

このサンプルでは、ホストは 1 つだけ指定されています (Outlook)。

```xml
<Host xsi:type="MailHost">
```
    
この要素内には、デスクトップ版 Outlook の構成の詳細が含まれています。

```xml
<DesktopFormFactor>
```
    
ボタン用のすべての JavaScript コードを含む HTML ファイルへの URL は、`FunctionFile` 要素で指定されています (`Resources` 要素が指定するリソース ID が使用される点に注意してください):

```xml
<FunctionFile resid="functionFile" />
```
    
マニフェストは、単一の拡張点を設定することにより、メッセージ フォームへのアクティブ化も制限します。

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
```
    
ボタンのプロパティは、`Control` 要素で指定されます。最も重要なのは、ボタンのクリック イベントが、`Action` 要素内にある `Translator.js` 内の `translate` 関数に接続されている点です。

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>translate</FunctionName>
</Action>
```
    
## 質問とコメント

- このサンプルの実行について問題がある場合は、[問題をログに記録](https://github.com/OfficeDev/Outlook-Add-in-Commands-Translator/issues)してください。
- Office アドイン開発全般の質問については、「[Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)」に投稿してください。質問やコメントには、必ず `office-addins` というタグを付けてください。

## その他のリソース

- [Outlook デベロッパー センター](https://dev.outlook.com)
- MSDN 上の[Office アドイン](https://msdn.microsoft.com/library/office/jj220060.aspx) ドキュメント
- [その他のアドイン サンプル](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## 著作権

Copyright (c) 2015 Microsoft.All rights reserved.


このプロジェクトでは、[Microsoft オープン ソース倫理規定](https://opensource.microsoft.com/codeofconduct/)が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。
