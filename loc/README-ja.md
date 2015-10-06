#iOS# 版 Office 365 Connect アプリ

[日本 (日本語)](/loc/README-ja.md) (日本語)

[![Office 365 Connect のサンプル](/readme-images/O365-iOS-Connect-video_play_icon.png)](https://youtu.be/3v__BnV61Rs "活用できるサンプルを確認するにはこちらをクリックしてください")

Office 365 への接続は、各 iOS アプリが Office 365 によって提供される豊富なデータとサービスの操作を開始するために必要な最初の手順です。このサンプルは、Office 365 に接続し、sendMail API を呼び出して、Office 365 の電子メール アカウントから電子メールを送信する方法を示します。このサンプルは、iOS アプリを Office 365 にすばやく接続する開始点として使用できます。Objective-C バージョンと Swift バージョンの両方でご利用いただけます。

**目次**

* [環境を設定する](#set-up-your-environment)
* [CocoaPods を使用して Office 365 iOS SDK をインポートする](#use-cocoapods-to-import-the-o365-ios-sdk)
* [Microsoft Azure にアプリを登録する](#register-your-app-with-microsoft-azure)
* [プロジェクトにクライアント ID とリダイレクト URI を取り込む](#get-the-client-id-and-redirect-uri-into-the-project)
* [目的のコード](#code-of-interest)
* [質問とコメント](#questions-and-comments)
* [トラブルシューティング](#troubleshooting)
* [その他の技術情報](#additional-resources)


<a name="set-up-your-environment"></a>
## 環境を設定する ##

iOS 版 Office 365 Connect アプリを実行するには、以下が必要です。


* 
            Apple 社の [Xcode](https://developer.apple.com/)。
* Office 365 アカウント。Office 365 アカウントは、[Office 365 開発者向けサイト](http://msdn.microsoft.com/library/office/fp179924.aspx)にサインアップすると取得できます。これにより、Office 365 のデータを対象とするアプリの作成に使用できる API にアクセスできるようになります。
* アプリケーションを登録する Microsoft Azure テナント。Azure Active Directory は、アプリケーションが認証と承認に使用する ID サービスを提供します。試用版のサブスクリプションを作成し、O365 アカウントに関連付ける方法の詳細については、「[Office 365 の開発環境を設定](https://msdn.microsoft.com/office/office365/howto/setup-development-environment)」および 「**新しい Azure サブスクリプションを作成し、Office 365 アカウントに関連付ける**」のセクションを参照してください。

  **重要事項**:Azure サブスクリプションが既ある場合、Office 365 アカウントにそのサブスクリプションをバインドする必要があります。この工程の詳細については、「[Office 365 の開発環境を設定する](https://msdn.microsoft.com/office/office365/howto/setup-development-environment)」 および「 **Office 365 アカウントを Azure AD と関連付けてアプリを作成および管理する**」のセクションを参照してください。


* 依存関係マネージャーとしての [CocoaPods](https://cocoapods.org/) のインストール。CocoaPods を使用すると、Office 365 と Azure Active Directory 認証ライブラリ (ADAL) の依存関係をプロジェクトに導入することができます。

Office 365 アカウント、および Office 365 開発者サイトにバインドされた Azure AD アカウントを取得したら、次の手順を実行する必要があります。

1. Microsoft Azure でアプリケーションを登録し、Office 365 Exchange Online の適切なアクセス許可を構成します。この方法については後で説明します。
2. CocoaPods をインストールし、これを使用して、プロジェクトに Office 365 と ADAL 認証の依存関係を取り込みます。この方法については後で説明します。
3. Azure アプリの登録固有の情報 (ClientID と RedirectUri) を、Office 365 Connect アプリに入力します。

<a name="use-cocoapods-to-import-the-o365-ios-sdk"></a>
## CocoaPods を使用して Office 365 iOS SDK をインポートする
注:依存関係マネージャーとして **CocoaPods** を初めて使用する場合は、これをインストールしてからプロジェクトで Office 365 iOS SDK の依存関係を取り込む必要があります。インストール済みの場合は、この手順を省略して「**プロジェクトに iOS 版 Office 365 SDK の依存関係を取り込む**」に移動しても構いません。

Mac の**ターミナル** アプリケーションから、次のコード行を入力します。

    sudo gem install cocoapods
    pod setup

インストールとセットアップが成功すると、「**ターミナルのセットアップが完了しました**」というメッセージが表示されます。CocoaPods とその使用法の詳細については、「[CocoaPods](https://cocoapods.org/)」を参照してください。


**プロジェクトに iOS 版 Office 365 SDK の依存関係を取り込む**
Office 365 iOS Connect アプリには、プロジェクトに Office 365 と ADAL のコンポーネント (pods) を取り込む podfile がすでに含まれています。podfile の場所は、サンプルの設定 (「Podfile」) に基づき、**objective-c** または **swift** のいずれかのフォルダーにあります。次の例は、このファイルの内容を示しています。

	target ‘test’ do

    pod 'ADALiOS', '~> 1.2.1'
    pod 'Office365/Outlook', '= 0.9.1'
    pod 'Office365/Discovery', '= 0.9.1'

	end


**Terminal** (プロジェクト フォルダーのルート) にあるプロジェクトのディレクトリに移動して、次のコマンドを実行する必要があります。


    pod install

注: これらの依存関係がプロジェクトに追加されたことを示す確認のメッセージを受信します。Podfile に構文エラーがあると、インストール コマンドを実行する際にエラーが発生します。

<a name="register-your-app-with-microsoft-azure"></a>
## Microsoft Azure にアプリを登録する
1.	Azure AD 資格情報を使用して、[Azure 管理ポータル](https://manage.windowsazure.com)にサインインします。
2.	左側のメニューで **[Active Directory]** をクリックしてから、Office 365 開発者向けサイトのディレクトリをクリックします。
3.	上部のメニューで、**[アプリケーション]** をクリックします。
4.	下部のメニューから、**[追加]** をクリックします。
5.	**[何を行いますか] ページ**で、**[所属組織が開発しているアプリケーションの追加]** をクリックします。
6.	**[アプリケーションについてお聞かせください]** ページで、アプリケーション名には「**O365-iOS-Connect**」を指定し、種類は「**ネイティブ クライアント アプリケーション**」を選択します。
7.	ページの右下隅にある矢印アイコンをクリックします。
8.	[アプリケーション情報] ページで、リダイレクト URI を指定します。この例では http://localhost/connect を指定します。続いて、ページの右下隅にあるチェック ボックスをクリックします。この値は、「**プロジェクトに ClientID と RedirectUri を取り込む**」セクションで使用するため覚えておいてください。
9.	アプリケーションが正常に追加されたら、アプリケーションの [クイック スタート] ページに移動します。ここで、上部のメニューにある [構成] をクリックします。
10.	**[他のアプリケーションへのアクセス許可]** で、次のアクセス許可を追加します。**Office 365 Exchange Online** アプリケーションを追加します。次に、アプリケーションを追加するために、右下隅のチェック ボックスをクリックします。最後に、**[他のアプリケーションへのアクセス許可]** セクションに戻り、**[委任されたアクセス許可]** から **[ユーザーとしてメールを送信]** を選択してください。
13.	**[構成]** ページで、**[クライアント ID]** に指定された値をコピーします。この値は、「**プロジェクトに ClientID と RedirectUri を取り込む**」セクションで使用するため覚えておいてください。
14.	下部のメニューで、**[保存]** をクリックします。

<a name="get-the-client-id-and-redirect-uri-into-the-project"></a>
## クライアント ID を取得して、URI をプロジェクトにリダイレクトする

最後に、前のセクション「**Microsoft Azure にアプリを登録する**」で記録したクライアント ID とリダイレクト URI を追加する必要があります。

**O365-iOS-Connect** プロジェクトのディレクトリを参照し、ワークスペース (O365-iOS-Connect.xcworkspace) を開きます。**AuthenticationManager** ファイルで、**ClientID** と **RedirectUri** の各値がファイルの一番上に追加されていることが分かります。ここで必要な値を指定します。

    // You will set your application's clientId and redirect URI. You get
    // these when you register your application in Azure AD.
    static NSString * const REDIRECT_URL_STRING = @"ENTER_REDIRECT_URI_HERE";
    static NSString * const CLIENT_ID           = @"ENTER_CLIENT_ID_HERE";
    static NSString * const AUTHORITY           = @"https://login.microsoftonline.com/common";



<a name="code-of-interest"></a>
## 目的のコード

**Azure AD での認証**

Azure AD での認証 (これにはアクセス トークンの取得と管理が含まれます) 用のコードは、AuthenticationManager にあります。


**Outlook サービス クライアント**

Outlook サービス クライアントを作成するコードは、Office365ClientFetcher にあります。このファイル内のコードは、Office 365 Exchange サービスに対する API の呼び出しを実行するために必要な Outlook サービス クライアント オブジェクトを作成します。クライアント要求は、認証コードを利用して、アプリケーションがユーザーの代理として操作するためのアクセスおよび更新トークンを取得できるようにします。アクセスおよび更新トークンはキャッシュされます。次回ユーザーがサービスへのアクセスを試みる際、アクセス トークンが発行されます。アクセス トークンの有効期限が切れている場合、クライアントは、新しいアクセス トークンを取得するために更新トークンを発行します。


**探索サービス**

Office 365 探索サービスを使用して Exchange サービスのエンドポイント/URL を取得するコードは、内部メソッド connectToOffice365 にあります。このメソッドは、SendMailViewController の viewDidLoad メソッドから呼び出されます。


**Office 365 SendMail スニペット**

メールを送信する操作のコードは、SendMailViewController の sendMailMessage メソッドにあります。


<a name="questions-and-comments"></a>
## 質問とコメント

iOS 版 Office 365 のサンプルについて、Microsoft にフィードバックをお寄せください。フィードバックは、このリポジトリの「[問題](https://github.com/OfficeDev/O365-iOS-Connect)」セクションに送信できます。<br>Office 365 の開発全般については、 「[スタックオーバーフロー](http://stackoverflow.com/questions/tagged/Office365+API)」に送信してください。質問には、必ず [Office365] および [API] のタグを付けてください。

<a name="troubleshooting"></a>
## トラブルシューティング

Xcode 7.0 のアップデートにより、iOS 9 を実行するシミュレーターやデバイス用に App Transport Security を使用できるようになりました。「[App Transport Security のテクニカル ノート](https://developer.apple.com/library/prerelease/ios/technotes/App-Transport-Security-Technote/)」を参照してください。

このサンプルでは、plist 内の次のドメインのために一時的な例外を作成しました:

- outlook.office365.com

これらの例外が含まれていないと、Xcode で iOS 9 シミュレーターにデプロイされたときに、このアプリで Office 365 API へのすべての呼び出しが失敗します。

<a name="additional-resources"></a>
## その他の技術情報

* [iOS 用 Office 365 コード スニペット](https://github.com/OfficeDev/O365-iOS-Snippets)
* [iOS 用 Office 365 プロファイル サンプル](https://github.com/OfficeDev/O365-iOS-Profile)
* [Email Peek - Office 365 を使用して構築された iOS アプリ](https://github.com/OfficeDev/O365-iOS-EmailPeek)
* [Office 365 API ドキュメント](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [ファイルの REST 操作のリファレンス](http://msdn.microsoft.com/office/office365/api/files-rest-operations)
* [予定表の REST 操作のリファレンス](http://msdn.microsoft.com/office/office365/api/calendar-rest-operations)
* [Office デベロッパー センター](http://dev.office.com/)
* [Office 365 API のサンプル コードとビデオ](https://msdn.microsoft.com/office/office365/howto/starter-projects-and-code-samples)



