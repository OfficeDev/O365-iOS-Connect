#iOS# 版 Office 365 Connect アプリ

Office 365 への接続は、各 iOS アプリが Office 365 によって提供される豊富なデータとサービスの操作を開始するために必要な最初の手順です。このサンプルは、Office 365 に接続し、sendMail API を呼び出して、Office 365 の電子メール アカウントから電子メールを送信する方法を示します。このサンプルは、iOS アプリを Office 365 にすばやく接続する開始点として使用できます。Objective-C と Swift の両方のバージョンが用意されています。

**目次**

* [環境を設定する](#set-up-your-environment)
* [CocoaPods を使用して Office 365 iOS SDK をインポートする](#use-cocoapods-to-import-the-o365-ios-sdk)
* [Microsoft Azure にアプリを登録する](#register-your-app-with-microsoft-azure)
* [プロジェクトにクライアント ID とリダイレクト URI を取り込む](#get-the-client-id-and-redirect-uri-into-the-project)
* [目的のコード](#code-of-interest)
* [質問とコメント](#questions-and-comments)
* [その他の技術情報](#additional-resources)


<a name="set-up-your-environment"></a>
## 環境を設定する ##

iOS 版 Office 365 Connect アプリを実行するには、以下が必要です。


* Apple 社の [Xcode](https://developer.apple.com/)。
* Office 365 アカウント。Office 365 アカウントは、[Office 365 開発者向けサイト](http://msdn.microsoft.com/ja-jp/library/office/fp179924.aspx)にサイン アップすると取得できます。これにより、Office 365 のデータを対象とするアプリの作成に使用できる API にアクセスできるようになります。
* アプリケーションを登録する Microsoft Azure テナント。Azure Active Directory は、アプリケーションが認証と承認に使用する ID サービスを提供します。ここでは、試用版サブスクリプションを取得できます。[Microsoft Azure](https://account.windowsazure.com/SignUp)。

  **重要事項**:Azure サブスクリプションが Office 365 テナントにバインドされていることを確認する必要があります。確認するには、Active Directory チームのブログ投稿「[複数の Windows Azure Active Directory を作成および管理する](http://blogs.technet.com/b/ad/archive/2013/11/08/creating-and-managing-multiple-windows-azure-active-directories.aspx)」を参照してください。「**新しいディレクトリを追加する**」セクションで、この方法について説明しています。また、詳細については、「[開発者向けサイトに Azure Active Directory へのアクセスを設定する](http://msdn.microsoft.com/en-us/office/office365/howto/setup-development-environment#bk_CreateAzureSubscription)」も参照してください。


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
Office 365 iOS Connect アプリには、プロジェクトに Office 365 と ADAL コンポーネント (pods) を取り込む podfile が既に含まれています。podfile がある場所は、サンプルの root ("Podfile") です。次の例は、ファイルの内容を示しています。

	target ‘test’ do

    pod 'ADALiOS', '~> 1.2.1'
    pod 'Office365/Outlook', '= 0.9.1'
    pod 'Office365/Discovery', '= 0.9.1'
    pod 'Office365/Files', '= 0.9.1'

	end


**Terminal** (プロジェクト フォルダーの root) にあるプロジェクトのディレクトリに移動して、次のコマンドを実行する必要があります。


    pod install

注:「これらの依存関係がプロジェクトに追加されました。今すぐ Xcode (**O365-iOS-Connect.xcworkspace**) でプロジェクトの代わりにワークスペースを開く必要があります」という確認のメッセージを受信する必要があります。Podfile で構文エラーが発生すると、インストール コマンドを実行する際にエラーが発生します。

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
10.	**[他のアプリケーションへのアクセス許可]** で、次のアクセス許可を追加します。**Office 365 Exchange Online アプリケーションを追加**してから、**ユーザーとしてメールを送信**のアクセス許可を選択します。
13.	**[構成]** ページで、**[クライアント ID]** に指定された値をコピーします。この値は、「**プロジェクトに ClientID と RedirectUri を取り込む**」セクションで使用するため覚えておいてください。
14.	下部のメニューで、**[保存]** をクリックします。

<a name="get-the-client-id-and-redirect-uri-into-the-project"></a>
## クライアント ID を取得して、URI をプロジェクトにリダイレクトする

最後に、前のセクション「**Microsoft Azure にアプリを登録する**」で記録したクライアント ID とリダイレクト URI を追加する必要があります。

**O365-iOS-Connect** プロジェクトのディレクトリを参照し、ワークスペース (O365-iOS-Connect.xcworkspace) を開きます。**AuthenticationManager.m** ファイルで、**ClientID** と **RedirectUri** の各値がファイルの一番上に追加されていることが分かります。ここで必要な値を指定します。

    // You will set your application's clientId and redirect URI. You get
    // these when you register your application in Azure AD.
    static NSString * const REDIRECT_URL_STRING = @"ENTER_REDIRECT_URI_HERE";
    static NSString * const CLIENT_ID           = @"ENTER_CLIENT_ID_HERE";
    static NSString * const AUTHORITY           = @"https://login.microsoftonline.com/common";



<a name="code-of-interest"></a>
## 目的のコード

**Azure AD での認証**

アクセス トークンの取得と管理が含まれる Azure AD での認証用のコードは、AuthenticationManager にあります。


**Outlook サービス クライアント**

Outlook サービス クライアントを作成するコードは、Office365ClientFetcher にあります。このファイル内のコードは、Office 365 Exchange サービスに対する API 呼び出しを実行するために必要な Outlook サービス クライアント オブジェクトを作成します。クライアント要求は、認証コードを利用して、アプリケーションがユーザーの代理として操作するためのアクセスおよび更新トークンを取得できるようにします。アクセスおよび更新トークンはキャッシュされます。次回ユーザーがサービスへのアクセスを試みる際、アクセス トークンが発行されます。アクセス トークンの有効期限が切れている場合、クライアントは、新しいアクセス トークンを取得するための更新トークンを発行します。


**探索サービス**

Office 365 探索サービスを使用して Exchange サービスのエンドポイント/URL を取得するコードは、内部メソッド connectToOffice365 にあります。このメソッドは、SendMailViewController の viewDidLoad メソッドから呼び出されます。


**Office 365 SendMail スニペット**

メールを送信する操作のコードは、SendMailViewController の sendMailMessage メソッドにあります。


<a name="questions-and-comments"></a>
## 質問とコメント

iOS 版 Office 365 のサンプルについて、Microsoft にフィードバックをお寄せください。フィードバックは、このリポジトリの「[問題](https://github.com/OfficeDev/O365-iOS-Connect)」セクションに送信できます。<br>Office 365 の開発全般については、「[スタック オーバーフロー](http://stackoverflow.com/questions/tagged/Office365+API)」に送信してください。質問には、必ず [Office365] および [API] のタグを付けてください。

<a name="additional-resources"></a>
## その他の技術情報

* [Office 365 API ドキュメント](http://msdn.microsoft.com/office/office365/howto/platform-development-overview)
* [ファイルの REST 操作のリファレンス](http://msdn.microsoft.com/en-us/office/office365/api/files-rest-operations)
* [予定表の REST 操作のリファレンス](http://msdn.microsoft.com/en-us/office/office365/api/calendar-rest-operations)
* [Office デベロッパー センター](http://dev.office.com/)
* [Windows 版 Office 365 API スタート プロジェクト](https://github.com/OfficeDev/Office-365-APIs-Starter-Project-for-Windows)
