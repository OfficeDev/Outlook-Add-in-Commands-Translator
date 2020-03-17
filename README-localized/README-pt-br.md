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

O suplemento Tradutor usa o modelo de comandos em suplementos do Outlook para adicionar um botão à Faixa de Opções no formulário de nova mensagem. O botão envia o texto selecionado do corpo da mensagem para um serviço Web de tradução para traduzir do português para o russo.

## Pré-requisitos

Para executar este exemplo, você precisará do seguinte:

- Um servidor Web para hospedar os arquivos de exemplo. O servidor deve ser capaz de aceitar solicitações protegidas por SSL (https) e ter um certificado SSL válido.
- Uma conta de e-mail do Office 365 **ou** do Outlook.com. Você pode [participar do Programa de Desenvolvedores do Office 365 e obter uma assinatura gratuita de um ano do Office 365](https://aka.ms/devprogramsignup).
- Uma chave de API para a [API Yandex Translate](https://translate.yandex.com/developers).
- O Outlook 2016, que faz parte da [Versão Prévia do Office 2016](https://products.office.com/en-us/office-2016-preview).

## Configurar e instalar o exemplo

1. Baixar ou bifurcar o repositório.
1. Abra o `TranslateHelper.js` em um editor de texto e substitua o valor `YOUR API KEY HERE` pelo valor da sua chave para a API Yandex Translate. Salve suas alterações.
1. Carregue os diretórios `AppCompose`, `Conteúdo`e `Imagens` em um diretório no servidor da Web.
1. Abra o `TranslateAppManifest.xml` em um editor de texto. Substitua todas as instâncias de `YOUR_WEB_SERVER` pela URL HTTPS do diretório em que você carregou os arquivos na etapa anterior. Salve suas alterações.
1. Em um navegador, faça logon na sua conta de e-mail no https://outlook.office365.com (para o Office 365) ou https://www.outlook.com (para o Outlook.com). Clique no ícone de engrenagem no canto superior direito e em **Gerenciar Suplementos**.
    
  ![O item de menu Gerenciar suplementos no https://www.outlook.com](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-manage-addins.PNG)
    
1. Na lista de suplementos, clique no ícone de **+** e escolha **Adicionar de um arquivo**.

  ![O item de menu Adicionar de arquivo na lista de suplementos](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/addin-list.PNG)

1. Clique em **Procurar** e navegue até o arquivo `TranslateAppManifest.xml` no computador de desenvolvimento. Clique em **Avançar**.

  ![Caixa de diálogo Adicionar suplemento de um arquivo](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/browse-manifest.PNG)

1. Na tela de confirmação, você verá um aviso informando que o suplemento não é da Office Store e ainda não foi verificado pela Microsoft. Clique em **Instalar**.
1. Você deverá ver uma mensagem de sucesso: **Você adicionou um suplemento do Outlook**. Clique em OK.

## Execução da amostra ##

1. Abra o Outlook 2016 e conecte-se à conta de e-mail na qual você instalou o suplemento.
1. Crie um novo e-mail. Observe que o suplemento colocou o botão **Traduzir** na faixa de opções de comando.

  ![O botão Traduzir em um novo formulário de e-mail no Outlook](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/new-mail.PNG)

1. Digite um texto em português no corpo. Selecione o texto e clique em **Traduzir**.

  ![Formulário de novo email com texto em português selecionado no corpo](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-selected.PNG)

1. Após alguns instantes, o texto selecionado será substituído pela tradução russa, e você verá a mensagem **Traduzido com êxito** na barra de informações.

  ![O texto traduzido em russo pelo suplemento](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-translated.PNG)

## Componentes principais do exemplo

- [```TranslateAppManifest.xml```](TranslateAppManifest.xml): o arquivo de manifesto do suplemento Tradutor.
- [```AppCompose/FunctionFile/Home.html```](AppCompose/FunctionFile/Home.html): um arquivo HTML vazio para carregar `Translator.js` em clientes que oferecem suporte aos comandos de suplementos.
- [```AppCompose/FunctionFile/Translator.js```](AppCompose/FunctionFile/Translator.js): o código chamado quando se clica nos botões de comando do suplemento.
- [```AppCompose/Home/Home.html```](AppCompose/Home/Home.html): o arquivo HTML carregado e exibido por clientes que não têm suporte para comandos de suplemento.
- [```AppCompose/Home/Home.js```](AppCompose/Home/Home.js): o código invocado por clientes que não têm suporte para comandos de suplemento.
- [```AppCompose/TranslateHelper.js```](AppCompose/TranslateHelper.js): código comum usado por `Translator.js` e `Home.js`.

## Como isso funciona?

A parte fundamental da amostra é a estrutura do arquivo de manifesto. O manifesto usa o mesmo esquema de versão 1.1 que qualquer suplemento do Office. No entanto, há uma nova seção do manifesto chamada `VersionOverrides`. Essa seção contém todas as informações que os clientes que oferecem suporte aos comandos do suplemento (no momento somente para o Outlook 2016) precisam para chamar o suplemento em um botão da faixa de opções. Colocando isso em uma seção completamente separada, o manifesto também pode incluir a marcação original para habilitar o suplemento a ser carregado em um painel de tarefas sob o modelo original do suplemento no modo de redação. Você pode ver isso em ação carregando o suplemento no Outlook 2013 ou no Outlook na Web.

### O suplemento Tradutor carregado no Outlook na Web ###

![O suplemento Tradutor carregado no Outlook na Web](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-on-web.PNG)

No elemento `VersionOverrides`, há dois elementos filho, `Recursos` e `Hosts`. O elemento `Recursos` contém informações sobre ícones, cadeias de caracteres e qual arquivo HTML carregar para o suplemento. A seção `Hosts` especifica como e quando o suplemento é carregado.

Neste exemplo, há apenas um host especificado (Outlook):

```xml
<Host xsi:type="MailHost">
```
    
Dentro desse elemento estão as especificações de configuração para a versão de área de trabalho do Outlook:

```xml
<DesktopFormFactor>
```
    
A URL para o arquivo HTML com todo o código JavaScript para o botão é especificado no elemento `Functionfile`, (observe que ele usa a ID do recurso especificado no elemento `Recursos`):

```xml
<FunctionFile resid="functionFile" />
```
    
O manifesto também limita a ativação do formulário nova mensagem ao configurar um único ponto de extensão:

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
```
    
As propriedades do botão são especificadas no elemento `Controle`. O mais importante é que o evento de clique do botão está conectado à função `Tradutor` em `Translator.js` dentro do elemento `Ação`:

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>translate</FunctionName>
</Action>
```
    
## Perguntas e comentários

- Se você tiver problemas para executar este exemplo, [relate um problema](https://github.com/OfficeDev/Outlook-Add-in-Commands-Translator/issues).
- Perguntas sobre o desenvolvimento de Suplementos do Office em geral devem ser postadas no [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Não deixe de marcar as perguntas ou comentários com `office-addins`.

## Recursos adicionais

- [Centro de Desenvolvimento do Outlook](https://dev.outlook.com)
- Documentação de [Suplementos do Office](https://msdn.microsoft.com/library/office/jj220060.aspx) no MSDN
- [Mais exemplos de Suplementos](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Direitos autorais

Copyright © 2015 Microsoft. Todos os direitos reservados.


Este projeto adotou o [Código de Conduta do Código Aberto da Microsoft](https://opensource.microsoft.com/codeofconduct/). Para saber mais, confira [Perguntas frequentes sobre o Código de Conduta](https://opensource.microsoft.com/codeofconduct/faq/) ou contate [opencode@microsoft.com](mailto:opencode@microsoft.com) se tiver outras dúvidas ou comentários.
