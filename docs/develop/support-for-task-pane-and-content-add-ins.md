---
title: Prise en charge de l’API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 86c5b0e19a767ae4715e48bef93465ab45a32e31
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506034"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Prise en charge de l’API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013


Vous pouvez utiliser l’[API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js) pour créer un complément du volet Office ou de contenu pour les applications hôtes d’Office 2013. Les méthodes et les objets pris en charge par les compléments du volet Office et de contenu sont classés de la manière suivante :


1. **Objets communs partagés avec d’autres compléments Office.** Parmi ces objets figurent [Office](https://docs.microsoft.com/javascript/api/office?view=office-js), [Context](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js) et [AsyncResult](https://docs.microsoft.com/javascript/api/office/office.asyncresult?view=office-js). L’objet **Office** est l’objet racine de l’interface API JavaScript pour Office. L’objet **Context** représente l’environnement d’exécution du complément. **Office** et **Context** sont les objets fondamentaux pour tout complément Office. L’objet **AsyncResult** représente les résultats d’une opération asynchrone, comme les données renvoyées vers la méthode **getSelectedDataAsync**, qui lit les éléments sélectionnés par un utilisateur dans un document.
    
2.  **The Document object.** The majority of the API available to content and task pane add-ins is exposed through the methods, properties, and events of the [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) object. A content or task pane add-in can use the [Office.context.document](https://docs.microsoft.com/javascript/api/office/office.context?view=office-js#document) property to access the **Document** object, and through it, can access the key members of the API for working with data in documents, such as the [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js) and [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js) objects, and the [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-), and [getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-) methods. The **Document** object also provides the [mode](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#mode) property for determining whether a document is read-only or in edit mode, the [url](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#url) property to get the URL of the current document, and access to the [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) object. The **Document** object also supports adding event handlers for the [SelectionChanged](https://docs.microsoft.com/javascript/api/office/office.documentselectionchangedeventargs?view=office-js) event, so you can detect when a user changes their selection in the document.
    
   Un complément de contenu ou du volet Office peut accéder à l’objet **Document** uniquement après le chargement de l’environnement d’exécution et du DOM, généralement dans le gestionnaire d’événements pour l’événement [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js). Pour plus d’informations sur le flux d’événements lors de l’initialisation d’un complément et sur la vérification du chargement correct du DOM et de l’environnement d’exécution, voir la page relative au [chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).
    
3.  **Objets pour l’utilisation de fonctionnalités spécifiques.** Pour travailler avec des fonctionnalités spécifiques de l’API, utilisez les méthodes et les objets suivants :
    
    - Les objets [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js), [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.binding?view=office-js) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.
    
    - Les objets [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js) et [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.
    
    - Les objets [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js) et [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js) pour créer une copie de l’intégralité du document, le diviser en blocs ou en « sections », puis lire ou transmettre les données dans ces sections.
    
    - L’objet [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) pour enregistrer des données personnalisées, telles que des préférences utilisateur et l’état du complément.
    

> [!IMPORTANT]
> Certains des membres d’API ne sont pas pris en charge dans toutes les applications Office pouvant héberger des compléments de contenu et du volet Office. Pour déterminer les membres pris en charge, voir les ressources suivantes :

Pour consulter un résumé de la prise en charge de l’API JavaScript pour Office dans les applications hôtes Office, reportez-vous à l’article [Présentation de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Lecture et écriture dans une sélection active

Vous pouvez lire ou écrire dans la sélection en cours de l’utilisateur dans un document, une feuille de calcul ou une présentation. Selon l’application hôte de votre complément, vous pouvez spécifier le type de structure de données à lire ou à écrire en tant que paramètre dans les méthodes [getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) et [setSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#setselecteddataasync-data--options--callback-) de l’objet [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js). Par exemple, vous pouvez indiquer n’importe quel type de données (HTML, données tabulaires, Office Open XML ou texte) pour Word, des données texte et tabulaires pour Excel et des données texte pour PowerPoint et Project. Vous pouvez également créer des gestionnaires d’événements pour détecter les modifications apportées à la sélection de l’utilisateur. L’exemple suivant obtient des données à partir de la sélection en tant que données texte à l’aide de la méthode **getSelectedDataAsync**.


```js
Office.context.document.getSelectedDataAsync(
    Office.CoercionType.Text, function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        }
        else {
            write('Selected data: ' + asyncResult.value);
        }
    });

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```

Pour plus d’informations et d’exemples, reportez-vous à l’article [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Liaison à une région d’un document ou d’une feuille de calcul

Vous pouvez utiliser les méthodes **getSelectedDataAsync** et **setSelectedDataAsync** pour lire ou écrire la sélection *en cours* de l’utilisateur dans un document, une feuille de calcul ou une présentation.

Vous pouvez ajouter une liaison à l’aide des méthodes [addFromNamedItemAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfrompromptasync-bindingtype--options--callback-) ou [addFromSelectionAsync](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js#addfromselectionasync-bindingtype--options--callback-) de l’objet [Bindings](https://docs.microsoft.com/javascript/api/office/office.bindings?view=office-js).

Pour plus d’informations et d’exemples, voir **Liaisons de régions dans un document ou une feuille de calcul**.



```js
Office.context.document.bindings.addFromSelectionAsync(
    Office.BindingType.Text, { id: 'myBinding' }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Added new binding with type: ' +
            asyncResult.value.type + ' and id: ' + asyncResult.value.id);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

Pour plus d’informations et d’exemples, reportez-vous à l’article [Liaisons de régions dans un document ou une feuille de calcul](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="getting-entire-documents"></a>Obtention de documents entiers

Si votre complément du volet Office s’exécute dans PowerPoint ou Word, vous pouvez utiliser les méthodes [Document.getFileAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfileasync-filetype--options--callback-), [File.getSliceAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#getsliceasync-sliceindex--callback-) et [File.closeAsync](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js#closeasync-callback-) pour obtenir la totalité d’une présentation ou d’un document.

Lorsque vous appelez **Document.getFileAsync**, vous obtenez une copie du document dans un objet [File](https://docs.microsoft.com/javascript/api/office/office.file?view=office-js). L’objet **File** donne accès au document en « blocs » représenté sous la forme d’objets [Slice](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js). Lorsque vous appelez **getFileAsync**, vous pouvez spécifier le type de fichier (texte ou format Open Office XML compressé) et la taille des secteurs (jusqu’à 4 Mo). Pour accéder au contenu de l’objet **File**, appelez **File.getSliceAsync** qui renvoie les données brutes dans la propriété [Slice.data](https://docs.microsoft.com/javascript/api/office/office.slice?view=office-js#data). Si vous avez spécifié un format compressé, vous obtiendrez les données du fichier sous la forme d’un tableau d’octets. Si vous transférez le fichier à un service web, vous pouvez transformer les données brutes compressées dans une chaîne codée en Base64 avant l’envoi. Enfin, lorsque vous avez obtenu les sections du fichier, utilisez la méthode **File.closeAsync** pour fermer le document.

Pour plus d’informations, reportez-vous à l’article expliquant comment [obtenir l’intégralité d’un document à partir d’un complément pour PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md). 


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Lecture et écriture des parties XML personnalisées d’un document Word

Grâce aux contrôles de contenu et au format de fichier Office Open XML, vous pouvez ajouter des parties XML personnalisées à un document Word et lier des éléments dans les parties XML aux contrôles de contenu de ce document. Lorsque vous ouvrez le document, Word lit et remplit automatiquement les contrôles de contenu liés avec les données des parties XML personnalisées. Les utilisateurs peuvent également écrire des données dans les contrôles de contenu. Lorsqu’ils enregistrent le document, les données des contrôles sont alors enregistrées dans les parties XML liées. Si votre complément du volet Office s’exécute dans Word, vous pouvez utiliser la propriété [Document.customXmlParts](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js.customxmlparts), ainsi que les objets [CustomXmlParts](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js), [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) et [CustomXmlNode](https://docs.microsoft.com/javascript/api/office/office.customxmlnode?view=office-js) pour lire et écrire des données de manière dynamique dans le document.

Les parties XML personnalisées peuvent être associées à des espaces de noms. Pour obtenir des données à partir des parties XML personnalisées dans un espace de noms, utilisez la méthode [CustomXmlParts.getByNamespaceAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbynamespaceasync-ns--options--callback-).

Vous pouvez également utiliser la [CustomXmlParts.getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-) pour accéder aux parties XML personnalisées par leur GUID. Après avoir obtenu une partie XML personnalisée, utilisez la méthode [CustomXmlPart.getXmlAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js#getxmlasync-options--callback-) pour obtenir les données XML.

Pour ajouter une partie XML personnalisée à un document, utilisez la propriété **Document.customXmlParts** afin d’obtenir les parties XML personnalisées qui sont dans le document et appelez la méthode [CustomXmlParts.addAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#addasync-xml--options--callback-).

Pour obtenir des informations détaillées sur l’utilisation de parties XML personnalisées avec un complément du volet Office, voir [Création de meilleurs compléments pour Word avec Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Persistance des paramètres de complément


Vous devez souvent enregistrer les données personnalisées pour votre complément, telles que les préférences d’un utilisateur ou l’état du complément, et accéder à ces données lors de la prochaine ouverture du complément. Vous pouvez utiliser des techniques de programmation web courantes pour enregistrer les données, comme les cookies de navigateur ou le stockage web HTML 5. Si votre complément est également exécuté dans Excel, PowerPoint ou Word, vous pouvez également utiliser les méthodes de l’objet [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js). Les données créées avec l’objet **Settings** sont stockées dans la feuille de calcul, la présentation ou le document dans lequel le complément a été inséré et enregistré. Ces données sont disponibles seulement pour le complément qui les a créées.

To avoid roundtrips to the server where the document is stored, data created with the **Settings** object is managed in memory at run time. Previously saved settings data is loaded into memory when the add-in is initialized, and changes to that data are only saved back to the document when you call the [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) method. Internally, the data is stored in a serialized JSON object as name/value pairs. You use the [get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-), [set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-), and [remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) methods of the **Settings** object, to read, write, and delete items from the in-memory copy of the data. The following line of code shows how to create a setting named `themeColor` and set its value to 'green'.




```js
Office.context.document.settings.set('themeColor', 'green');
```

Étant donné que les données de paramètres créées ou supprimées avec les méthodes **set** et **remove** agissent sur une copie en mémoire des données, vous devez appeler **saveAsync** pour rendre persistantes les modifications apportées aux données de paramètres dans le document utilisé par votre complément.

Pour plus de détails sur l’utilisation des données personnalisées à l’aide de l’objet **Settings**, reportez-vous à l’article [Conservation de l’état et des paramètres des compléments](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Lecture des propriétés d’un document de projet

Si votre complément de volet Office s’exécute dans Project, vous pouvez lire les données de certains champs, ressources et champs de tâche du projet actif. Pour ce faire, vous pouvez utiliser les méthodes et les événements de l’objet [ProjectDocument](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js), qui étend l’objet **Document** pour fournir des fonctionnalités supplémentaires propres à Project.

Pour des exemples de lecture de données Project, reportez-vous à l’article [Créer votre premier complément du volet Office pour Project 2013 à l’aide d’un éditeur de texte](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


## <a name="permissions-model-and-governance"></a>Modèle d’autorisations et gouvernance

Votre complément utilise l’élément **Permissions** dans son manifeste pour demander l’autorisation d’accéder au niveau de fonctionnalité qu’il exige à partir de l’interface API JavaScript pour Office. Par exemple, si votre complément nécessite un accès en lecture/écriture pour le document, son manifeste doit spécifier `ReadWriteDocument` en tant que valeur de texte dans l’élément **Permissions**. Étant donné que les autorisations ont pour objectif de protéger la vie privée et la sécurité de l’utilisateur, en tant que meilleures pratiques, nous vous recommandons de demander le niveau d’autorisation minimal requis pour ses fonctionnalités. L’exemple suivant illustre la demande de l’autorisation **ReadDocument** dans le manifeste d’un complément du volet Office.


```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Pour plus d’informations, consultez l’article [Demande d’autorisations pour l’utilisation d’API dans les compléments de contenu et du volet Office](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).


## <a name="see-also"></a>Voir aussi

- [API JavaScript pour Office](https://docs.microsoft.com/office/dev/add-ins/reference/javascript-api-for-office?view=office-js)
- [Référence de schéma pour les manifestes des compléments Office](../develop/add-in-manifests.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
    
