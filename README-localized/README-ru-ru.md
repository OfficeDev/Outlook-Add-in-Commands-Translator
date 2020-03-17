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

Надстройка "Переводчик" добавляет кнопку на ленту в форме создания сообщения, используя модель команд для надстроек Outlook. Кнопка отправляет выделенный текст из тела сообщения в веб-службу переводчика для перевода с английского на русский.

## Условия

Для работы с этим образцом необходимо выполнить указанные ниже действия.

- Веб-сервер для размещения примеров файлов. Сервер должен иметь возможность принимать запросы, защищенные SSL (https), и иметь действительный сертификат SSL.
- Учетная запись электронной почты Office 365 **или** учетная запись электронной почты Outlook.com. Вы можете [присоединиться к Программе разработчика Office 365 и получить бесплатную годовую подписку на Office 365](https://aka.ms/devprogramsignup).
- Ключ API для [API Переводчика Яндекса](https://translate.yandex.com/developers).
- Outlook 2016, который является частью [Office 2016 Preview](https://products.office.com/en-us/office-2016-preview).

## Настройка и установка образца

1. Загрузите или разветвите репозиторий.
1. В текстовом редакторе откройте `TranslateHelper.js` и замените `ключ API здесь` значение с помощью ключа API для Яндекс преобразования. Сохраните изменения.
1. Добавьте `AppCompose`, `контента`и `изображения` каталогов в каталог на веб-сервере.
1. Откройте `TranslateAppManifest.xml` в текстовом редакторе. Замените все экземпляры `YOUR_WEB_SERVER` на HTTPS URL-адрес каталога, куда вы загрузили файлы на предыдущем шаге. Сохраните изменения.
1. Войдите в свою учетную запись электронной почты с помощью браузера по адресу https://outlook.office365.com (для Office 365) или https://www.outlook.com (для Outlook.com). Нажмите на значок шестеренки в правом верхнем углу, затем нажмите **Управление надстройками**.
    
  ![Пункт меню Управление надстройками на https://www.outlook.com](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-manage-addins.PNG)
    
1. В списке надстроек щелкните значок **+** и выберите **добавить из файла**.

  ![Пункт "Добавить из файла" в списке надстроек](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/addin-list.PNG)

1. Нажмите кнопку **Обзор** и найдите файл `TranslateAppManifest.xml` на компьютере разработчика. Нажмите **Далее**.

  ![Диалоговое окно "Добавление надстройки из файла"](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/browse-manifest.PNG)

1. На экране подтверждения вы увидите предупреждение о том, что надстройка не из Office Store и не была проверена Microsoft. Нажмите **Установить**.
1. Должно отобразиться сообщение об успешном выполнении: Удаление**вы добавили надстройку для Outlook**. надстройки Outlook Нажмите кнопку ОК.

## Запуск приложения ##

1. Откройте Outlook 2016 и подключитесь к учетной записи электронной почты, в которой вы установили надстройку.
1. Создайте новое сообщение. Обратите внимание, что надстройка поместила кнопку **Перевести** на командную ленту.

  ![Кнопка «Перевести» в новой почтовой форме в Outlook](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/new-mail.PNG)

1. Введите текст на английском языке в тексте. Выделите текст, затем нажмите **Перевести**.

  ![Новая почтовая форма с английским текстом, выбранным в теле](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-selected.PNG)

1. Через некоторое время выделенный текст будет заменен русским переводом, и вы должны увидеть сообщение **Успешно переведенный** на информационной панели.

  ![Текст переведен на русский язык надстройкой](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-translated.PNG)

## Ключевые компоненты примера

- [```TranslateAppManifest.xml```](TranslateAppManifest.xml): Файл манифеста для надстройки Translator.
- [```AppCompose/FunctionFile/Home.html```](AppCompose/FunctionFile/Home.html): Пустой HTML-файл для загрузки `Translator.js` для клиентов, которые поддерживают команды надстроек.
- [```AppCompose/FunctionFile/Translator.js```](AppCompose/FunctionFile/Translator.js): Код, который вызывается при нажатии кнопки команды надстройки.
- [```AppCompose/Home/Home.html```](AppCompose/Home/Home.html): Файл HTML, который загружается и отображается клиентами, которые не поддерживают команды надстроек.
- [```AppCompose/Home/Home.js```](AppCompose/Home/Home.js): Код, который вызывается клиентами, которые не поддерживают команды надстроек.
- [```AppCompose/TranslateHelper.js```](AppCompose/TranslateHelper.js): Общий код, используемый как `Translator.js`, так и `Home.js`.

## Как все это работает?

Ключевой частью примера является структура файла манифеста. В манифесте используется та же схема версии 1.1, что и в манифесте любой надстройки Office. Однако есть новый раздел манифеста под названием `VersionOverrides`. В этом разделе содержится вся информация, необходимая клиентам, которые поддерживают команды надстроек (в настоящее время только Outlook 2016), для вызова надстройки с помощью кнопки ленты. Поместив это в совершенно отдельный раздел, манифест может также включать исходную разметку, чтобы позволить загрузке надстройки на панели задач в исходной модели надстройки режима компоновки. Вы можете увидеть это в действии, загрузив надстройку в Outlook 2013 или Outlook в Интернете.

### Надстройка переводчика загружена в Outlook в Интернете ###

![Надстройка переводчика загружена в Outlook в Интернете](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-on-web.PNG)

В элементе `VersionOverrides` есть два дочерних элемента: `ресурсы` и `хосты`. Элемент `Ресурсы` содержит информацию о значках, строках и HTML-файле, который нужно загрузить для надстройки. В разделе `Хосты` указывается, как и когда загружается надстройка.

В этом примере указан только один хост (Outlook):

```xml
<Host xsi:type="MailHost">
```
    
Внутри этого элемента находятся особенности конфигурации для настольной версии Outlook:

```xml
<DesktopFormFactor>
```
    
URL-адрес HTML-файла со всем кодом JavaScript для кнопки указывается в элементе `FunctionFile` (обратите внимание, что он использует идентификатор ресурса, указанный в элементе `Ресурсы`):

```xml
<FunctionFile resid="functionFile" />
```
    
Манифест также ограничивает активацию новой формой сообщения, устанавливая одну точку расширения:

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
```
    
Свойства кнопки указываются в элементе `управления`. Самое главное, что событие нажатия кнопки связано с функцией `перевода` в `Translator.js` внутри элемента `Action`:

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>translate</FunctionName>
</Action>
```
    
## Вопросы и комментарии

- Если у вас возникли проблемы с запуском этого примера, [сообщите о неполадке](https://github.com/OfficeDev/Outlook-Add-in-Commands-Translator/issues).
- Вопросы о разработке надстроек Office в целом следует размещать в [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Убедитесь в том, что ваши вопросы и комментарии помечены `Office надстройки`.

## Дополнительные ресурсы

- [Центр разработки для Outlook](https://dev.outlook.com)
- [надстройки Office](https://msdn.microsoft.com/library/office/jj220060.aspx) документации в MSDN
- [Дополнительные примеры надстроек](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Авторские права

(c) Корпорация Майкрософт (Microsoft Corporation), 2015. Все права защищены.


Этот проект соответствует [Правилам поведения разработчиков открытого кода Майкрософт](https://opensource.microsoft.com/codeofconduct/). Дополнительные сведения см. в разделе [часто задаваемых вопросов о правилах поведения](https://opensource.microsoft.com/codeofconduct/faq/). Если у вас возникли вопросы или замечания, напишите нам по адресу [opencode@microsoft.com](mailto:opencode@microsoft.com).
