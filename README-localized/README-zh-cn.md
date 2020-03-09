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
# Microsoft Teams“待办事项列表”选项卡应用示例
[![生成状态](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo.svg?branch=master)](https://travis-ci.org/OfficeDev/microsoft-teams-sample-todo)

这是 [Microsoft Teams 选项卡应用](https://aka.ms/microsoftteamstabsplatform)的示例。本示例旨在说明将现有 Web 应用转换为 Microsoft Teams 选项卡应用是一个非常简单的过程。现有的 web 应用 [**TodoMVC for React**](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react)，作为基本的任务管理器，它与个人 Outlook 任务相集成。只进行了少量修改，该 web 视图可作为选项卡应用添加至频道。查看“[‘之前’和‘之后’分支的代码差异](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after)”，了解更改的内容。

> **注意：**这不是团队协作应用程序的实例。所显示的任务属于用户的个人帐户，不属于共享团队帐户。

**有关 Microsoft Teams 开发体验的详细信息，请参阅 Microsoft Teams [开发人员文档](https://msdn.microsoft.com/en-us/microsoft-teams/index)。**

## 先决条件

1. 拥有[可以访问 Microsoft Teams 的 Office 365 账户](https://msdn.microsoft.com/en-us/microsoft-teams/setup)。
2. 此示例使用 [Node.js](https://nodejs.org) 生成。如果尚未安装，请下载并安装建议的版本。

## 运行应用

### 托管选项卡应用 "configuration page" 和 "content page"

如果选修卡配置和内容页面是 web 应用，必须进行托管。  

首先，我们将准备项目：

1. 克隆存储库。
2. 在存储库子目录中打开命令行。
3. 使用命令行运行 `npm install`（作为 Node.js 一部分）。

若要手动托管应用：
* 在存储库子目录中，运行 `npm start` 以启动`webpack-dev-server` 来启用开发人员应用程序。
*注意，开发人员清单设置使用 localhost，因此无需附加配置。

或者，创建可在内部环境中托管的可部署应用：
* 在存储库子目录中，运行 `npm run build` 来生成可部署版本。
* 修改生产清单中的 config.url，以指向托管位置。[请参阅下方内容](#registering-an-application-to-authenticate-with-microsoft)

### 将选项卡添加到 Microsoft Teams

1. 为此示例下载“[选项卡应用开发包](https://github.com/OfficeDev/microsoft-teams-sample-todo/tree/master/package/todo.dev.zip)”。
2. 根据需要，新建团队进行测试。单击左侧窗格底部的“**创建团队**”。
3. 从左侧窗格中选择团队，选择 **... （更多选项）**，然后选择“**查看团队**”。
4. 选择“**开发人员（预览版）**”选项卡，随后选择“**上传**”。
5. 导航至第 1 步中下载的压缩文件，并选中。
6. 转至团队中的任一频道。单击现有选项卡右侧的 "+"。
7. 从显示的库中选择选项卡。
8. 接受许可提示。
9. 如果需要，使用 Office 365 工作/学校帐户登录。注意，代码将尝试执行静默式身份验证。
10. 验证身份验证信息。
11. 点击“保存”以将选项卡添加到频道。

> **注意：**若要重新上传具有相同 `id` 的已更新的程序包，请单击该选项卡表末的“替换”图标。请勿再次单击“上传”：Microsoft Teams 将提示选项卡已存在。

> 建议拥有多个配置，每个环境一个。压缩文件的名称可以是`todo.dev.zip`、`todo.prod.zip` 等，但压缩文件必须包含具有唯一 `id` 的 `manifest.json`。参见“[创建选项卡清单](https://msdn.microsoft.com/en-us/microsoft-teams/createpackage)”。

## 代码演练

如果`主`分支显示最新的示例状态，查看下列两项之间的下列[代码差异](https://github.com/OfficeDev/microsoft-teams-sample-todo/compare/before...after)：

* [`之前`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/before)：初始应用

* [`之后`](https://github.com/OfficeDev/microsoft-teams-sample-todo/commit/after)：与 Microsoft Teams 团队集成后的应用。

逐步完成此步骤：

1. 我们添加了新 `config.html` 和 `config.tsx` 页面，它们负责应用程序，允许用户在首次登录期间操纵任意设置、执行单一登录身份验证等。团队管理员配置应用程序/设置，需要此操作。参见“[创建配置页面](https://msdn.microsoft.com/en-us/microsoft-teams/createconfigpage)”。

2. 我们新增了相同的 `config.html` 文件至 `webpack.common.js`配置中，因此可在运行时期间插入正确的捆绑包。

3. 我们添加了 `MicrosoftTeams.js`引用至 index.html中，并为 Typescript intellisense 添加了`MicrosoftTeams.d.ts`。

4. 新增了`manifest.dev.json`、`manifest.prod.json` 和大小为 *44x44* 和 *88x88* 的两个图标。记住在上传至 Microsoft Teams 的压缩文件中，将之重命名为 `manifest.json`。

5. 我们添加了部分样式至 `app.css`中。

6. 最后修改了 outlook.tasks.ts 中的“`身份验证`”，以依赖于 'useMicrosoftTeamsAuth'，它是此示例中所引用试用版 OfficeHelpers 的新功能。

### 调用身份验证对话框

用户添加选项卡时，配置页面显示（config.html）。在这种情况中，代码可能验证用户的身份。

身份验证利用了最新版本 (>0.4.0) [office-js-helpers](https://github.com/OfficeDev/office-js-helpers)库的 Teams 专用新功能。该帮助程序功能将尝试静默进行身份验证，如果无法进行，将调用 Microsoft Teams 专用身份验证对话框。有关 Microsoft Teams 完整的身份验证流程详情，请参阅 Microsoft Teams [开发人员文档](https://msdn.microsoft.com/en-us/microsoft-teams/index)中的[验证用户](https://msdn.microsoft.com/en-us/microsoft-teams/auth)。

在 outlook.tasks.ts:
```typescript
 return this.authenticator.authenticate('Microsoft')
    .then(token => this._token = token)
    .catch(error => {
        Utilities.log(error);
        throw new Error('Failed to login using your Microsoft Account');
    });
```

### 处理“保存”事件

登录成功后，用户将保存选项卡至频道中。下列代码启用保存按钮，并设定 SaveHandler，它将保存选项卡中显示的内容（在此情况中，仅项目的 index.html）。

在 config.tsx:
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

## 使用的技术

此示例使用下列堆叠：

1. [`React by Facebook`](https://facebook.github.io/react/) 作为用户界面框架。
2. [`TypeScript`](https://www.typescriptlang.org/) 作为转换编译器。
4. [`TodoMVC`](http://todomvc.com/examples/typescript-react/#/) base for TodoMVC 功能。
5. [`Webpack`](https://webpack.github.io/) 作为生成工具。

## 其他资源

### 注册应用程序，以与 Microsoft 进行身份验证
注意，如果使用上述本地开发人员选项，将不需要身份验证，但是如果选择在自有环境中托管此选项卡，必须注册应用程序以进行身份验证。

1. 转到 [Microsoft 应用程序注册门户](https://apps.dev.microsoft.com)。
2. 使用 Office 365 工作/学校帐户登录。请勿使用 Microsoft 个人账户。
2. 添加新应用程序。
2. 记下新的`应用程序 ID`。
2. 单击“`添加平台`”并选择 `Web`。
3. 选中“`允许隐式流`”，并将重定向 URL 配置为 `https//< mywebsite >/config.html`。

有关托管选项卡页面的更多详细信息，参见 [Microsoft Teams 入门示例自述文件](https://github.com/OfficeDev/microsoft-teams-sample-get-started#host-tab-pages-over-https)。

>**注意：**默认情况下，组织应允许创建新应用程序。如果不允许，可获取开发人员测试租户，一年的免费 Office 365 开发人员试用订阅。[方法如下](https://msdn.microsoft.com/en-us/microsoft-teams/setup)。


## 制作人员

此项目基于 TodoMVC Typescript - 反应模板位于[这里](https://github.com/tastejs/todomvc/tree/gh-pages/examples/typescript-react)。

## 参与

请阅读“[参与](contributing.md)”，详细了解我们的行为守则和提交拉取请求的流程。

此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。

## 版本控制

我们使用 [SemVer](http://semver.org/) 进行版本控制。对于可用版本，参见“[此存储库上的标记](https://github.com/officedev/microsoft-teams-sample-todo/tags)”。

## 许可证

此项目使用 MIT 许可证授权 - 参见“[许可证”](LICENSE)文件了解详细信息
