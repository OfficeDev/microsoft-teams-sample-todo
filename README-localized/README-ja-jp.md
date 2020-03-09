---
page_type: sample
products:
- office-teams
- office-outlook
- office-365
languages:
- typescript
- nodejs
extensions:
  contentType: samples
  technologies:
  - Tabs
  createdDate: 10/19/2016 12:06:51 PM
---
# Microsoft Teams 'Todo List' サンプル タブ アプリ
[![ビルドの状態](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo.svg?branch=master)](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo)

これは、[Microsoft Teams 用のタブ アプリ](https://aka.ms/microsoftteamstabsplatform)の例です。このサンプルの目的は、既存の Web アプリをいかに簡単に Microsoft Teams タブ アプリに変換できるかを示すことです。既存の Web アプリである [**TodoMVC for React**](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react) は、個人の Outlook タスクと統合する基本的なタスク マネージャーを提供します。わずかな変更を加えるだけで、この Web ビューをタブ アプリとしてチャネルに追加できます。['before' ブランチと 'after' ブランチのコードの違い](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after)を見て、どのような変更が行われたかを確認してください。

> **注:**これは、チーム コラボレーション アプリの現実的な例ではありません。表示されるタスクは、共有チーム アカウントではなく、ユーザーの個人アカウントに属します。

**Microsoft Teams のエクスペリエンス開発の詳細については、Microsoft Teams [開発者向けドキュメント](https://msdn.microsoft.com/en-us/microsoft-teams/index)をご覧ください。**

## 前提条件

1. [Microsoft Teams へのアクセス権を持つ Office 365 アカウント](https://msdn.microsoft.com/en-us/microsoft-teams/setup)。
2. このサンプルは [Node.js](https://nodejs.org) を使用して構築されています。推奨バージョンをまだお持ちではない場合は、ダウンロードしてインストールしてください。

## アプリの実行

### タブ アプリの "構成ページ" と "コンテンツ ページ" をホストする

タブ構成およびコンテンツ ページは Web アプリであるため、ホストする必要があります。  

まず、プロジェクトを準備します。

1. リポジトリを複製します。
2. repo サブディレクトリでコマンド ラインを開きます。
3. コマンド ラインから `npm install` (Node.js の一部として含まれています) を実行します。

アプリをローカルでホストするには:
* repo サブディレクトリで、`npm start` を実行して `webpack-dev-server` を起動し、dev アプリが機能するようにします。
* dev マニフェストは localhost を使用するように設定されているため、追加の構成は必要ありません。

必要に応じて、独自の環境でホストできる展開可能なアプリをビルドするには、次のようにします。
* repo サブディレクトリで、`npm run build` を実行して展開可能なビルドを生成します。
* ホスティング場所を指すように、prod マニフェストの config.url を変更します。[以下をご覧ください](#registering-an-application-to-authenticate-with-microsoft)

### Microsoft Teams にタブを追加します

1. このサンプルの[タブ アプリ dev パッケージ](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/master/package/todo.dev.zip) zip ファイルをダウンロードします。
2. 必要に応じて、テスト用の新しいチームを作成します。左側のパネルの一番下にある [**チームの作成**] をクリックします。
3. 左側のパネルからチームを選択し、**... (その他のオプション)** を選択してから、[**チームの表示**]を選択します。
4. [**開発者 (プレビュー)**] タブを選択し、[**アップロード**] を選択します。
5. 上記の手順 1 からダウンロードした zip ファイルに移動して、選択します。
6. チーム内の任意のチャネルに移動します。既存のタブの右側にある '+' をクリックします。
7. 表示されるギャラリーからタブを選択します。
8. 同意プロンプトを受け入れます。
9. 必要に応じて、Office 365 の職場または学校のアカウントを使用してサインインします。コードは、可能であればサイレント認証を試行することに注意してください。
10. 認証情報を検証します。
11. [保存] をクリックして、タブをチャネルに追加します。

> **注:**同じ `id` を使用して更新されたパッケージを再アップロードするには、タブのテーブル行の最後にある [置換] アイコンをクリックします。もう一度 [アップロード] をクリックしないでください。Microsoft Teams は、タブが既に存在していると告げます。

> 環境ごとに 1 つずつ、複数の設定を持つことをお勧めします。zip ファイルの名前は、`todo.dev.zip`、`todo.prod.zip` などの任意の名前にすることができますが、zip には `manifest.json` (一意の `id` を持つ) が含まれている必要があります。「[Creating a manifest for your tab (タブ用のマニフェストの作成)](https://msdn.microsoft.com/en-us/microsoft-teams/createpackage)」をご覧ください。

## コードの段階的な説明

`master` ブランチにはサンプルの最新の状態が表示されますが、次の[コードの違い](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after)を見てください。

* [`before`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/before): 初期のアプリ

* [`after`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/after): Microsoft Teams と統合した後のアプリ。

手順を段階的に説明していきます。

1. 新しい `config.html` および `config.tsx` ページを追加しました。このページは、最初の起動時にユーザーが設定を操作したり、シングル サインオン認証を実行したりできるようにするアプリケーションを担当します。これは、チーム管理者がアプリケーション/設定を構成できるようにするために必要です。「[Create the configuration page (構成ページを作成する)](https://msdn.microsoft.com/en-us/microsoft-teams/createconfigpage)」を参照してください。

2. 同じ `config.html` ファイルを `webpack.common.js` 構成に追加して、実行中に適切なバンドルを挿入できるようにしました。

3. index.html に `MicrosoftTeams.js` への参照を追加し、Typescript intellisense に `MicrosoftTeams.d.ts` を追加しました。

4. `manifest.dev.json`、`manifest.prod.json`、およびサイズが *44x44* と *88x88* の 2 つのロゴを追加しました。Microsoft Teams にアップロードする zip ファイルで、これらの名前を `manifest.json` に変更することを忘れないでください。

5. `app.css` にいくつかのスタイルを追加しました。

6. 最後に、代わりにこのサンプルで参照されているベータ版の OfficeHelpers の新機能である 'useMicrosoftTeamsAuth' に依存するように、outlook.tasks.ts の`認証`を変更しました。

### [認証] ダイアログの呼び出し

ユーザーがタブを追加すると、構成ページ (config.html) が表示されます。この場合、可能であればコードはユーザーを認証します。

認証は、[office-js-helpers](https://github.com/OfficeDev/office-js-helpers) ライブラリの最新バージョン (0.4.0 より上) の新しい Teams 固有の機能を活用します。このヘルパー関数は、サイレント認証を試みますが、認証できない場合は、Microsoft Teams 固有の認証ダイアログを呼び出します。Microsoft Teams の完全な認証プロセスの詳細については、「[Authenticating a user (ユーザーの認証)](https://msdn.microsoft.com/en-us/microsoft-teams/auth)」 (Microsoft Teams [開発者向けドキュメント](https://msdn.microsoft.com/en-us/microsoft-teams/index)内) を参照してください。

outlook.tasks.ts 内:
```typescript
 return this.authenticator.authenticate('Microsoft')
    .then(token => this._token = token)
    .catch(error => {
        Utilities.log(error);
        throw new Error('Failed to login using your Microsoft Account');
    });
```

### ' 保存 ' イベントの処理

サインインに成功すると、ユーザーはタブをチャネルに保存します。次のコードは、[保存] ボタンを有効にし、SaveHandler を設定します。SaveHandler は、タブに表示するコンテンツ (この場合はプロジェクトの index.html のみ) を格納します。

config.tsx 内:
```typescript
    initialize({ groupId, upn}) {
        this.setState({ groupId, upn });
        console.log(this.state);
        /** Enable the Save button  */
        microsoftTeams.settings.setValidityState(true);
        /** Register the save handler */
        microsoftTeams.settings.registerOnSaveHandler(saveEvent => {
            /** Store Tab content settings */
            microsoftTeams.settings.setSettings({
                contentUrl: `${location.origin}/index.html`,
                suggestedDisplayName: "My Tasks",
                websiteUrl: `${location.origin}/index.html`
            });
            saveEvent.notifySuccess();
        });
    }
```

## 使用されるテクノロジ

このサンプルでは、次のスタックを使用します。

1. UI フレームワークとして [`React by Facebook`](https://facebook.github.io/react/)。
2. トランスパイラーとしての [`TypeScript`](https://www.typescriptlang.org/)。
4. TodoMVC 機能のベースとして[`TodoMVC`](http://todomvc.com/examples/typescript-react/#/)。
5. ビルド ツールとしての [`Webpack`](https://webpack.github.io/)。

## 追加情報

### Microsoft で認証するためのアプリケーションの登録
上記のローカル Dev オプションを使用する場合はこれは必要ありませんが、独自の環境でこのタブをホストすることを選択した場合は、認証するためにアプリケーションを登録する必要があります。

1. [Microsoft アプリケーション登録ポータル](https://apps.dev.microsoft.com)に移動します。
2. Office 365 の職場または学校のアカウントを使用してサインインします。個人用の Microsoft アカウントは使用しないでください。
2. 新しいアプリを追加します。
2. 新しい`アプリケーション ID` をメモします。
2. [`プラットフォームの追加`] をクリックし、[`Web`] を選びます。
3. [`暗黙的フローを許可する`] をオンにして、リダイレクト URL を `https://<mywebsite>/config.html` に設定します。

独自のタブ ページのホストの詳細については、[Microsoft Teams 'はじめに' サンプル README](https://github.com/OfficeDev/microsoft-teams-sample-get-started#host-tab-pages-over-https)を参照してください。

>**注:**既定では、組織は新しいアプリの作成を許可する必要があります。そうでない場合は、開発者テスト テナント、Office 365 Developer の 1 年間の試用版サブスクリプションを無料で入手できます。[これを行うには、次の操作を実行します](https://msdn.microsoft.com/en-us/microsoft-teams/setup)。


## クレジット

このプロジェクトは、[ここ](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react)にある TodoMVC Typescript - React テンプレートに基づいています。

## 貢献

行動規範の詳細とプル リクエストを送信するプロセスについては、「[貢献](contributing.md)」をお読みください。

このプロジェクトでは、[Microsoft Open Source Code of Conduct (Microsoft オープン ソース倫理規定)](https://opensource.microsoft.com/codeofconduct/) が採用されています。詳細については、「[倫理規定の FAQ](https://opensource.microsoft.com/codeofconduct/faq/)」を参照してください。また、その他の質問やコメントがあれば、[opencode@microsoft.com](mailto:opencode@microsoft.com) までお問い合わせください。

## バージョン管理

バージョン管理には [SemVer](http://semver.org/) を使用しています。使用可能なバージョンについては、[このリポジトリのタグ](https://github.com/officedev/microsoft-teams-sample-todo/tags)を参照してください。

## ライセンス

このプロジェクトは、MIT ライセンスの下でライセンスされています。詳細については、[ライセンス](LICENSE) ファイルを参照してください。
