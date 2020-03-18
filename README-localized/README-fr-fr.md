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

Le complément Translator utilise le modèle de commandes pour que les compléments Outlook ajoutent un bouton au ruban du formulaire de nouveau message. Le bouton envoie le texte sélectionné du corps du message à un service web de traducteur pour traduire de l’anglais vers le russe.

## Conditions préalables

Pour exécuter cet exemple, vous devez disposer des éléments ci-après :

- Un serveur web pour héberger les fichiers de l'exemple. Le serveur doit pouvoir accepter des demandes protégées par SSL (https) et disposer d’un certificat SSL valide.
- Un compte de messagerie Office 365 **ou** un compte de messagerie Outlook.com. Vous pouvez [participer au programme pour les développeurs Office 365 et obtenir un abonnement gratuit d’un an à Office 365](https://aka.ms/devprogramsignup).
- Clé API pour l’[API Yandex translate](https://translate.yandex.com/developers).
- Outlook 2016, qui fait partie de la [préversion Office 2016](https://products.office.com/en-us/office-2016-preview).

## Configuration et installation de l'exemple

1. Téléchargez ou dérivez par le référentiel.
1. Ouvrez `TranslateHelper.js` dans un éditeur de texte et remplacez la valeur `votre clé API ici` par votre clé API pour l’API Yandex translate. Enregistrez vos modifications.
1. Téléchargez les annuaires `AppCompose`, `Content` et `Images` vers un répertoire sur votre serveur Web.
1. Ouvrez `TranslateAppManifest.xml` dans un éditeur de texte. Remplacez toutes les instances de `VOTRE_SERVEUR_WEB` par l’URL de HTTPS du répertoire dans lequel vous avez téléchargé les fichiers au cours de l’étape précédente. Enregistrez vos modifications.
1. Connectez-vous à l'aide d'un navigateur à votre compte de messagerie sur https://outlook.office365.com (pour Office 365) ou https://www.outlook.com (pour Outlook.com). Cliquez sur l’icône d’engrenage dans le coin supérieur droit et choisissez **Gérer les compléments**.
    
  ![L'élément de menu Gérer des compléments sur https://outlook.com](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-manage-addins.PNG)
    
1. Dans la liste de compléments, cliquez sur l’icône **+**, puis sélectionnez **Ajouter à partir d’un fichier**.

  ![Élément de menu Ajouter à partir du fichier dans la liste des compléments](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/addin-list.PNG)

1. Cliquex sur **Parcourir** et accédez au fichier `TranslateAppManifest.xml` sur votre ordinateur de développement. Cliquez sur **Suivant**.

  ![Ajouter un complément à partir de la boîte de dialogue Fichier](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/browse-manifest.PNG)

1. L’écran de confirmation affiche un message d'alerte indiquant que le complément ne provient pas du Store Office et n’a pas été vérifié par Microsoft. Cliquez sur **Installer**.
1. Un message de réussite doit s'afficher : **Vous avez ajouté un complément pour Outlook**. Cliquez sur OK.

## Exécution de l’exemple ##

1. Ouvrez Outlook 2016 et connectez-vous au compte de messagerie sur lequel vous avez installé le complément.
1. Créez un courrier électronique. Veuillez noter que le complément a installé un nouveau bouton **Traduire** sur le ruban de commandes.

  ![Bouton Traduire sur un nouveau formulaire de courrier dans Outlook](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/new-mail.PNG)

1. Tapez du texte en anglais dans le corps du texte. Sélectionnez le texte, puis cliquez sur **Traduire**.

  ![Formulaire nouveau message avec le texte en anglais sélectionné dans le corps](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-selected.PNG)

1. Après quelques instants, le texte sélectionné est remplacé par la traduction russe, et le message **traduit correctement** apparaît dans la barre d’informations.

  ![Texte traduit en russe par le complément](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/text-translated.PNG)

## Composants clés de l’exemple

- [```TranslateAppManifest.xml```](TranslateAppManifest.xml): Fichier manifeste pour le complément Traduire.
- [```AppCompose/FunctionFile/Home.html```](AppCompose/FunctionFile/Home.html): Fichier HTML vide pour charger `Translator.js` pour les clients prenant en charge les commandes de complément.
- [```AppCompose/FunctionFile/Translator.js```](AppCompose/FunctionFile/Translator.js): Code appelé lorsqu'un utilisateur clique sur les boutons de commande du complément.
- [```AppCompose/Home/Home.html```](AppCompose/Home/Home.html): Fichier HTML chargé et affiché par les clients qui ne prennent pas en charge les commandes de complément.
- [```AppCompose/Home/Home.js```](AppCompose/Home/Home.js): Code qui est appelé par les clients qui ne prennent pas en charge les commandes de complément.
- [```AppCompose/TranslateHelper.js```](AppCompose/TranslateHelper.js): Code commun utilisé par `Translator.js` et `Home.js`.

## Comment fonctionne tout cela ?

L'élément clé de l’exemple est la structure du fichier manifeste. Le manifeste utilise le même schéma de version 1.1 que tout manifeste de complément Office. Cependant, une nouvelle section du manifeste appelée `VersionOverrides` existe. Cette section comprend toutes les informations que les clients prenant en charge les commandes de complément (actuellement seulement Outlook 2016) doivent suivre pour appeler le complément à partir d’un bouton du ruban. En plaçant ceci dans une section totalement distincte, le manifeste peut également inclure la balise d’origine permettant le chargement du complément dans un volet de tâches sous le modèle original du complément en mode composition. Vous pouvez afficher ceci en action en chargeant le complément dans Outlook 2013 ou Outlook sur le web.

### Le complément Traducteur chargé dans Outlook sur le web ###

![Le complément Traducteur chargé dans Outlook sur le web](https://raw.githubusercontent.com/OfficeDev/Outlook-Add-in-Commands-Translator/master/readme-images/outlook-on-web.PNG)

Dans l'élément `VersionOverrides`, il existe deux éléments enfants : `Ressources`et `Hôtes`. L’élément `Ressources` se compose des informations sur les icônes, les chaînes et le fichier HTML à charger pour le complément. La section `Hôtes` spécifie la modalité et la période de chargement du complément.

Dans cet exemple, un seul hôte est mentionné (Outlook) :

```xml
<Host xsi:type="MailHost">
```
    
Les détails de la configuration de la version de bureau d’Outlook sont inclus dans cet élément :

```xml
<DesktopFormFactor>
```
    
L’URL du fichier HTML contenant la totalité du code JavaScript pour le bouton est spécifiée dans l'élément `FunctionFile` (veuillez noter qu’il utilise l’ID de ressource spécifié dans l’élément `Ressources`) :

```xml
<FunctionFile resid="functionFile" />
```
    
Le manifeste limite également l’activation au nouveau formulaire de message en définissant un seul point d’extension :

```xml
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
```
    
Les propriétés du bouton sont spécifiées dans l’élément `contrôle`. Plus important, l’événement clic du bouton est connecté à la fonction `traduire` dans `Translator.js` à l’intérieur de l’élément `Action` :

```xml
<Action xsi:type="ExecuteFunction">
    <FunctionName>translate</FunctionName>
</Action>
```
    
## Questions et commentaires

- Si vous rencontrez des difficultés pour exécuter cet exemple, veuillez [consigner un problème](https://github.com/OfficeDev/Outlook-Add-in-Commands-Translator/issues).
- Si vous avez des questions générales sur le développement de compléments Office, envoyez-les sur [Stack Overflow](http://stackoverflow.com/questions/tagged/office-addins). Posez vos questions ou envoyez vos commentaires en incluant la balise `office-addins`.

## Ressources supplémentaires

- [Centre des développeurs Outlook](https://dev.outlook.com)
- Documentation pour [Compléments Office](https://msdn.microsoft.com/library/office/jj220060.aspx) sur MSDN.
- [Autres exemples de compléments](https://github.com/OfficeDev?utf8=%E2%9C%93&query=-Add-in)

## Copyright

Copyright (c) 2015 Microsoft. Tous droits réservés.


Ce projet a adopté le [code de conduite Open Source de Microsoft](https://opensource.microsoft.com/codeofconduct/). Pour en savoir plus, reportez-vous à la [FAQ relative au code de conduite](https://opensource.microsoft.com/codeofconduct/faq/) ou contactez [opencode@microsoft.com](mailto:opencode@microsoft.com) pour toute question ou tout commentaire.
