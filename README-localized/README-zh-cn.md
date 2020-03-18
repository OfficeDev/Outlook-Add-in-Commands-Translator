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

翻译加载项将命令模型用于 Outlook 加载项以向新邮件窗体中的功能区添加按钮。此按钮将邮件正文中的选定文本发送到翻译工具 Web 服务，将英语文本翻译为俄语。

## 先决条件

若要运行此示例，需要具备以下条件：

- 用于托管示例文件的 Web 服务器。服务器必须能够接受受 SSL 保护的请求 (https)，并且具备有效的 SSL 证书。
- Office 365 电子邮件帐户**或** Outlook.com 电子邮件帐户。你可以[参加 Office 365 开发人员计划并获取为期 1 年的免费 Office 365 订阅](https://aka.ms/devprogramsignup)。
- [Yandex Translate API](https://translate.yandex.com/developers) 的 API密钥。
- Outlook 2016（[Office 2016 预览版的一部分](https://products.office.com/en-us/office-2016-preview)）。

## 配置和安装示例

1. 下载或为存储库创建分支。
1. 在文本编辑器中打开 `TranslateHelper.js`，并将 `YOUR API KEY HERE` 值替换为 Yandex Translate API 的 API 密钥。保存所做的更改。
1. 将 `AppCompose`、`Content` 和 `Images` 目录上传到 Web 服务器上的目录。
1. 在文本编辑器中打开 `TranslateAppManifest.xml`。将 `YOUR_WEB_SERVER` 的所有实例替换为在上一步中上传的文件所在目录的 HTTPS URL。保存所做的更改。
1. 使用浏览器在 https://outlook.office365.com（对于 Office 365）或 https://www.outlook.com（对于 Outlook.com）上登录电子邮件帐户。单击右上角的齿轮图标，然后选择“**管理加载项**”。
    
  ![https://www.outlook.com 上的“管理加载项”菜单项](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-manage-addins.PNG)
    
1. 在加载项列表中，单击 **+** 图标并选择“**从文件添加**”。

  ![加载项列表中的“从文件添加”菜单项](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/addin-list.PNG)

1. 单击“**浏览**”并浏览到开发计算机上的 `TranslateAppManifest.xml` 文件。单击“**下一步**”。

  ![“从文件添加加载项”对话框](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/browse-manifest.PNG)

1. 在确认屏幕上，你将看到一条警告，指出该加载项不是来自 Office 应用商店，并且尚未得到 Microsoft 的验证。单击“**安装**”。
1. 应看到一条成功消息：**你已添加一个 Outlook 相关加载项**。单击“确定”。

## 运行示例 ##

1. 打开 Outlook 2016 并连接到安装了该加载项的电子邮件帐户。
1. 新建一封电子邮件。请注意，该加载项已在命令功能区上放置了“**翻译**”按钮。

  ![Outlook 中的“新建邮件”窗体上的“翻译”按钮](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/new-mail.PNG)

1. 在正文中键入一些英文文本。选择文本，然后单击“**翻译**”。

  ![在正文中选择了英文文本的“新建邮件”窗体](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-selected.PNG)

1. 片刻之后，所选文本将被俄语翻译替换，并且你应该会在信息栏中看到消息“**翻译成功**”。

  ![加载项将文本翻译成俄语](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-translated.PNG)

## 示例主要组件

- [```TranslateAppManifest.xml```](TranslateAppManifest.xml)：翻译工具加载项的清单文件。
- [```AppCompose/FunctionFile/Home.html```](AppCompose/FunctionFile/Home.html)：空 HTML 文件，用于为支持加载项命令的客户端加载 `Translator.js`。
- [```AppCompose/FunctionFile/Translator.js```](AppCompose/FunctionFile/Translator.js)：单击加载项命令按钮时调用的代码。
- [```AppCompose/Home/Home.html```](AppCompose/Home/Home.html)：由不支持加载项命令的客户端加载和显示的 HTML 文件。
- [```AppCompose/Home/Home.js```](AppCompose/Home/Home.js)：由不支持加载项命令的客户端调用的代码。
- [```AppCompose/TranslateHelper.js```](AppCompose/TranslateHelper.js)：供 `Translator.js` 和 `Home.js` 使用的常用代码。

## 它是如何工作的？

该示例的关键部分是清单文件的结构。该清单使用与任何 Office 加载项清单相同的版本 1.1 架构。然而，该清单中有一个称为 `VersionOverrides` 的新部分。此部分包含支持加载项命令（当前仅限 Outlook 2016）的客户端从功能区按钮调用加载项所需的所有信息。通过将其置于完全独立的部分中，该清单还可包含原始标记，以允许将加载项加载到原始撰写模式加载项模型下的任务窗格中。你可以通过在 Outlook 2013 或 Outlook 网页版中加载该加载项来了解工作方式。

### Outlook 网页版中加载的翻译加载项 ###

![Outlook 网页版中加载的翻译工具加载项](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-on-web.PNG)

在 `VersionOverrides` 元素中，有两个子元素，即`资源`和`主机`。`资源`元素包含有关图标、字符串以及要为加载项加载的 HTML 文件的信息。`主机`部分指定加载加载项的方式和时间。

在此示例中，仅指定了一个主机 (Outlook)：

```xml
<Host xsi:type="MailHost">
```
    
此元素包含 Outlook 桌面版本的配置详细信息：

```xml
<DesktopFormFactor>
```
    
HTML 文件的 URL，该文件包含在 `FunctionFile` 元素（请注意，它使用在`资源`元素中指定的资源 ID）中指定的按钮的所有 JavaScript 代码：

```xml
<FunctionFile resid="functionFile" />
```
    
该清单还通过设置单个扩展点来限制对“新建邮件”窗体的激活：

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
```
    
已在`控件`元素中指定按钮的属性。最重要的是，按钮单击事件与`操作`元素内的 `Translator.js` 中的 `translate` 函数相连：

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>translate</FunctionName>
</Action>
```
    
## 问题和意见

- 如果你在运行此示例时遇到任何问题，请[记录问题](https://github.com/OfficeDev/Outlook-Add-in-Commands-Translator/issues)。
- 与 Office 外接程序开发相关的问题一般应发布到 [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins)。确保你的问题或意见标记有 `Office 加载项`。

## 其他资源

- [Outlook 开发人员中心](https://dev.outlook.com)
- MSDN 上的 [Office 外接程序](https://msdn.microsoft.com/library/office/jj220060.aspx)文档
- [更多加载项示例](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## 版权信息

版权所有 (c) 2015 Microsoft。保留所有权利。


此项目已采用 [Microsoft 开放源代码行为准则](https://opensource.microsoft.com/codeofconduct/)。有关详细信息，请参阅[行为准则 FAQ](https://opensource.microsoft.com/codeofconduct/faq/)。如有其他任何问题或意见，也可联系 [opencode@microsoft.com](mailto:opencode@microsoft.com)。
