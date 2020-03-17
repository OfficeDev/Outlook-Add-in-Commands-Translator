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

El complemento “Translator” usa el modelo de comandos para complementos de Outlook para agregar un botón a la cinta de opciones en el formulario de mensaje nuevo. El botón envía el texto seleccionado en el cuerpo del mensaje a un servicio web de traducción para traducir del inglés al ruso.

## Requisitos previos

Para ejecutar este ejemplo, necesitará lo siguiente:

- Un servidor web para hospedar los archivos de ejemplo. El servidor debe poder aceptar solicitudes protegidas por SSL (https) y tener un certificado SSL válido.
- Una cuenta de correo electrónico de Office 365 **o** una cuenta de correo electrónico de Outlook.com. Puede [participar en el programa para desarrolladores Office 365 y obtener una suscripción gratuita durante 1 año a Office 365](https://aka.ms/devprogramsignup).
- Una clave de API para la [API de Yandex Translate](https://translate.yandex.com/developers).
- Outlook 2016, que forma parte de la [versión preliminar de Office 2016](https://products.office.com/en-us/office-2016-preview).

## Configurar e instalar el ejemplo

1. Descargue o bifurque el repositorio.
1. Abra `TranslateHelper.js` en un editor de texto y reemplace el valor `YOUR API KEY HERE` por la clave de API para la API de Yandex Translate. Guarde los cambios.
1. Cargue los directorios `AppCompose`, `Content` e `Images` en un directorio de su servidor web.
1. Abra `TranslateAppManifest.xml` en un editor de texto. Reemplace todas las instancias de `YOUR_WEB_SERVER` por la URL HTTPS del directorio donde se encuentran los archivos cargados en el paso anterior. Guarde los cambios.
1. Inicie sesión en su cuenta de correo electrónico con un explorador en https://outlook.office365.com (para Office 365) o https://www.outlook.com (para Outlook.com). Haga clic en el icono de engranaje en la esquina superior derecha y después haga clic en **Administrar complementos**.
    
  ![Elemento de menú Administrar complementos en https://www.outlook.com](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-manage-addins.PNG)
    
1. En la lista de complementos, haga clic en el icono **+** y elija **Agregar desde un archivo**.

  ![Elemento de menú Agregar desde archivo de la lista de complementos](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/addin-list.PNG)

1. Haga clic en **Examinar** y vaya al archivo `TranslateAppManifest.xml` del equipo de desarrollo. Haga clic en **Siguiente**.

  ![Diálogo Agregar complemento desde archivo](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/browse-manifest.PNG)

1. En la pantalla de confirmación, verá una advertencia que indica que el complemento no es de la Tienda Office y no está comprobado por Microsoft. Haga clic en **Instalar**.
1. Debería aparecer un mensaje que indica que todo ha ido bien: **Ha agregado un complemento para Outlook**. Haga clic en Aceptar.

## Ejecutar el ejemplo ##

1. Abra Outlook 2016 y conéctese a la cuenta de correo electrónico en la que instaló el complemento.
1. Cree un nuevo correo electrónico. Observe que el complemento ha colocado el botón **Traducir** en la cinta de comandos.

  ![Botón Traducir en un formulario de nuevo correo de Outlook](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/new-mail.PNG)

1. Escriba texto en inglés en el cuerpo. Seleccione el texto y haga clic en **Traducir**.

  ![Formulario de nuevo correo con el texto en inglés seleccionado en el cuerpo](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-selected.PNG)

1. Tras unos instantes, se reemplazará el texto seleccionado por la traducción al ruso y aparecerá el mensaje **Traducido correctamente** en la barra de información.

  ![Texto traducido al ruso por el complemento](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-translated.PNG)

## Componentes clave del ejemplo

- [```TranslateAppManifest.xml```](TranslateAppManifest.xml): archivo de manifiesto del complemento Translator.
- [```AppCompose/FunctionFile/Home.html```](AppCompose/FunctionFile/Home.html): archivo HTML vacío para cargar `Translator.js` para los clientes que admiten los comandos de complemento.
- [```AppCompose/FunctionFile/Translator.js```](AppCompose/FunctionFile/Translator.js): el código que se llama cuando se hace clic en el botón del comando de complemento.
- [```AppCompose/Home/Home.html```](AppCompose/Home/Home.html): el archivo HTML que cargan y muestran los clientes que no admiten los comandos de complemento.
- [```AppCompose/Home/Home.js```](AppCompose/Home/Home.js): el código que llaman los clientes que no admiten los comandos de complemento.
- [```AppCompose/TranslateHelper.js```](AppCompose/TranslateHelper.js): Código común usado por `Translator.js` y `Home.js`.

## ¿Cómo funciona?

La parte esencial del ejemplo es la estructura del archivo de manifiesto. El manifiesto usa el mismo esquema de la versión 1.1 que el manifiesto de cualquier complemento de Office. Sin embargo, hay una nueva sección del manifiesto llamada `VersionOverrides`. En esta sección se incluye toda la información que los clientes que admiten los comandos de complemento (actualmente solo Outlook 2016) necesitan para llamar al complemento desde un botón de la cinta de opciones. Al colocar esto en una sección completamente independiente, el manifiesto también puede incluir el formato original para permitir que el complemento se cargue en un panel de tareas en el modelo de complementos original del modo de redacción. Puede cargar el complemento en Outlook 2013 o en Outlook en la Web para ver cómo funciona.

### Complemento Translator cargado en Outlook en la Web ###

![Complemento Translator cargado en Outlook en la Web](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-on-web.PNG)

En el elemento `VersionOverrides`, hay dos elementos secundarios: `Resources` y `Hosts`. El elemento `Resources` contiene información sobre iconos y cadenas, así como sobre el archivo HTML que hay que cargar para el complemento. La sección `Hosts` especifica cómo y cuándo se carga el complemento.

En este ejemplo, solo se especifica un host (Outlook):

```xml
<Host xsi:type="MailHost">
```
    
En este elemento, se especifican las opciones de configuración para la versión de escritorio de Outlook:

```xml
<DesktopFormFactor>
```
    
La URL del archivo HTML que contiene todo el código JavaScript para el botón se especifica en el elemento `FunctionFile` (tenga en cuenta que usa el Id. de recurso que se especifica en el elemento `Resources`):

```xml
<FunctionFile resid="functionFile" />
```
    
El manifiesto también limita la activación al formulario de mensaje nuevo al establecer un único punto de extensión:

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
```
    
Las propiedades del botón se especifican en el elemento `Control`. Lo más importante es que el evento de clic del botón está conectado a la función `Translate` de `Translator.js` dentro del elemento `Action`:

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>translate</FunctionName>
</Action>
```
    
## Preguntas y comentarios

- Si tiene algún problema para ejecutar este ejemplo, [registre un problema](https://github.com/OfficeDev/Outlook-Add-in-Commands-Translator/issues).
- Las preguntas sobre el desarrollo de complementos para Office en general deben enviarse a [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Asegúrese de que sus preguntas o comentarios se etiquetan con `office-addins`.

## Recursos adicionales

- [Centro para desarrolladores de Outlook](https://dev.outlook.com)
- Documentación de [complementos de Office](https://msdn.microsoft.com/library/office/jj220060.aspx) en MSDN
- [Más complementos de ejemplo](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Derechos de autor

Copyright (c) 2015 Microsoft. Todos los derechos reservados.


Este proyecto ha adoptado el [Código de conducta de código abierto de Microsoft](https://opensource.microsoft.com/codeofconduct/). Para obtener más información, vea [Preguntas frecuentes sobre el código de conducta](https://opensource.microsoft.com/codeofconduct/faq/) o póngase en contacto con [opencode@microsoft.com](mailto:opencode@microsoft.com) si tiene otras preguntas o comentarios.
