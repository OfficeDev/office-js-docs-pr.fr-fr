---
title: Prise en charge de l’API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: aeb75d2dbdcae5636bbe81820b9387bd840120c7
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449308"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Prise en charge de l’interface API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013


Vous pouvez utiliser l’[API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office) pour créer un complément du volet Office ou de contenu pour les applications hôtes d’Office 2013. Les méthodes et les objets pris en charge par les compléments du volet Office et de contenu sont classés de la manière suivante :


1. **Objets communs partagés avec d’autres compléments Office.** Parmi ces objets figurent [Office](/javascript/api/office), [Context](/javascript/api/office/office.context) et [AsyncResult](/javascript/api/office/office.asyncresult). L’objet **Office** est l’objet racine de l’interface API JavaScript pour Office. L’objet **Context** représente l’environnement d’exécution du complément. **Office** et **Context** sont les objets fondamentaux pour tout complément Office. L’objet **AsyncResult** représente les résultats d’une opération asynchrone, comme les données renvoyées vers la méthode **getSelectedDataAsync**, qui lit les éléments sélectionnés par un utilisateur dans un document.

2. **Objet document.** La majorité des éléments de l’API disponibles pour les compléments de contenu et du volet Office sont exposés via les méthodes, propriétés et événements de l’objet [Document](/javascript/api/office/office.document). Un complément de contenu ou de volet de tâches peut utiliser la propriété [Office. Context. document](/javascript/api/office/office.context#document) pour accéder à l'objet **document** , et via ce dernier, peut accéder aux membres clés de l'API pour utiliser des données dans des documents, telles que les [liaisons](/javascript/api/office/office.bindings) et [ ](/javascript/api/office/office.customxmlparts)Les objets CustomXmlParts, ainsi que les méthodes [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-), [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-)et [getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-) . L’objet **Document** fournit également la propriété [mode](/javascript/api/office/office.document#mode) permettant de déterminer si un document est en mode lecture seule ou modification, la propriété [url](/javascript/api/office/office.document#url) pour obtenir l’URL du document actuel et accéder à l’objet [Settings](/javascript/api/office/office.settings). L'objet **document** prend également en charge l'ajout de gestionnaires d'événements pour l'événement [SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) , afin que vous puissiez détecter lorsqu'un utilisateur modifie sa sélection dans le document.

   Un complément de contenu ou du volet Office peut accéder à l’objet **Document** uniquement après le chargement de l’environnement d’exécution et du DOM, généralement dans le gestionnaire d’événements pour l’événement [Office.initialize](/javascript/api/office). Pour plus d’informations sur le flux d’événements lors de l’initialisation d’un complément et sur la vérification du chargement correct du DOM et de l’environnement d’exécution, voir la page relative au [chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

3. **Objets pour l’utilisation de fonctionnalités spécifiques.** Pour travailler avec des fonctionnalités spécifiques de l’API, utilisez les méthodes et les objets suivants :

    - Les objets [CustomXmlParts](/javascript/api/office/office.bindings), [CustomXmlPart](/javascript/api/office/office.binding) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.

    - Les objets [CustomXmlParts](/javascript/api/office/office.customxmlparts) et [CustomXmlPart](/javascript/api/office/office.customxmlpart) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.

    - Les objets [File](/javascript/api/office/office.file) et [Slice](/javascript/api/office/office.slice) pour créer une copie de l’intégralité du document, le diviser en blocs ou en « sections », puis lire ou transmettre les données dans ces sections.

    - L’objet [Settings](/javascript/api/office/office.settings) pour enregistrer des données personnalisées, telles que des préférences utilisateur et l’état du complément.


> [!IMPORTANT]
> Certains des membres d’API ne sont pas pris en charge dans toutes les applications Office pouvant héberger des compléments de contenu et du volet Office. Pour déterminer les membres pris en charge, voir les ressources suivantes :

Pour consulter un résumé de la prise en charge de l’API JavaScript pour Office dans les applications hôtes Office, reportez-vous à l’article [Présentation de l’API JavaScript pour Office](understanding-the-javascript-api-for-office.md).


## <a name="reading-and-writing-to-an-active-selection"></a>Reading and writing to an active selection

Vous pouvez lire ou écrire dans la sélection en cours de l’utilisateur dans un document, une feuille de calcul ou une présentation. Selon l’application hôte de votre complément, vous pouvez spécifier le type de structure de données à lire ou à écrire en tant que paramètre dans les méthodes [getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) et [setSelectedDataAsync](/javascript/api/office/office.document#setselecteddataasync-data--options--callback-) de l’objet [Document](/javascript/api/office/office.document). Par exemple, vous pouvez indiquer n’importe quel type de données (HTML, données tabulaires, Office Open XML ou texte) pour Word, des données texte et tabulaires pour Excel et des données texte pour PowerPoint et Project. Vous pouvez également créer des gestionnaires d’événements pour détecter les modifications apportées à la sélection de l’utilisateur. L’exemple suivant obtient des données à partir de la sélection en tant que données texte à l’aide de la méthode **getSelectedDataAsync**.


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

For more details and examples, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).


## <a name="binding-to-a-region-in-a-document-or-spreadsheet"></a>Binding to a region in a document or spreadsheet

Vous pouvez utiliser les méthodes **getSelectedDataAsync** et **setSelectedDataAsync** pour lire ou écrire la sélection *en cours* de l’utilisateur dans un document, une feuille de calcul ou une présentation.

Vous pouvez ajouter une liaison à l’aide des méthodes [addFromNamedItemAsync](/javascript/api/office/office.bindings#addfromnameditemasync-itemname--bindingtype--options--callback-), [addFromPromptAsync](/javascript/api/office/office.bindings#addfrompromptasync-bindingtype--options--callback-) ou [addFromSelectionAsync](/javascript/api/office/office.bindings#addfromselectionasync-bindingtype--options--callback-) de l’objet [Bindings](/javascript/api/office/office.bindings).

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

For more details and examples, see [Bind to regions in a document or spreadsheet](bind-to-regions-in-a-document-or-spreadsheet.md).


## <a name="getting-entire-documents"></a>Getting entire documents

Si votre complément du volet Office s’exécute dans PowerPoint ou Word, vous pouvez utiliser les méthodes [Document.getFileAsync](/javascript/api/office/office.document#getfileasync-filetype--options--callback-), [File.getSliceAsync](/javascript/api/office/office.file#getsliceasync-sliceindex--callback-) et [File.closeAsync](/javascript/api/office/office.file#closeasync-callback-) pour obtenir la totalité d’une présentation ou d’un document.

Lorsque vous appelez **Document.getFileAsync**, vous obtenez une copie du document dans un objet [File](/javascript/api/office/office.file). L’objet **File** donne accès au document en « blocs » représenté sous la forme d’objets [Slice](/javascript/api/office/office.slice). Lorsque vous appelez **getFileAsync**, vous pouvez spécifier le type de fichier (texte ou format Open Office XML compressé) et la taille des secteurs (jusqu’à 4 Mo). Pour accéder au contenu de l’objet **File**, appelez **File.getSliceAsync** qui renvoie les données brutes dans la propriété [Slice.data](/javascript/api/office/office.slice#data). Si vous avez spécifié un format compressé, vous obtiendrez les données du fichier sous la forme d’un tableau d’octets. Si vous transférez le fichier à un service web, vous pouvez transformer les données brutes compressées dans une chaîne codée en Base64 avant l’envoi. Enfin, lorsque vous avez obtenu les sections du fichier, utilisez la méthode **File.closeAsync** pour fermer le document.

For more details, see how to [get the whole document from an add-in for PowerPoint or Word](../word/get-the-whole-document-from-an-add-in-for-word.md).


## <a name="reading-and-writing-custom-xml-parts-of-a-word-document"></a>Reading and writing custom XML parts of a Word document

Grâce aux contrôles de contenu et au format de fichier Office Open XML, vous pouvez ajouter des parties XML personnalisées à un document Word et lier des éléments dans les parties XML aux contrôles de contenu de ce document. Lorsque vous ouvrez le document, Word lit et remplit automatiquement les contrôles de contenu liés avec les données des parties XML personnalisées. Les utilisateurs peuvent également écrire des données dans les contrôles de contenu. Lorsqu’ils enregistrent le document, les données des contrôles sont alors enregistrées dans les parties XML liées. Si votre complément du volet Office s’exécute dans Word, vous pouvez utiliser la propriété [Document.customXmlParts](/javascript/api/office/office.document#customxmlparts), ainsi que les objets [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) et [CustomXmlNode](/javascript/api/office/office.customxmlnode) pour lire et écrire des données de manière dynamique dans le document.

Les parties XML personnalisées peuvent être associées à des espaces de noms. Pour obtenir des données à partir des parties XML personnalisées dans un espace de noms, utilisez la méthode [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getbynamespaceasync-ns--options--callback-).

Vous pouvez également utiliser la [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) pour accéder aux parties XML personnalisées par leur GUID. Après avoir obtenu une partie XML personnalisée, utilisez la méthode [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getxmlasync-options--callback-) pour obtenir les données XML.

Pour ajouter une partie XML personnalisée à un document, utilisez la propriété **Document.customXmlParts** afin d’obtenir les parties XML personnalisées qui sont dans le document et appelez la méthode [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#addasync-xml--options--callback-).

Pour obtenir des informations détaillées sur l’utilisation de parties XML personnalisées avec un complément du volet Office, voir [Création de meilleurs compléments pour Word avec Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).


## <a name="persisting-add-in-settings"></a>Persistance des paramètres de complément


Vous devez souvent enregistrer les données personnalisées pour votre complément, telles que les préférences d’un utilisateur ou l’état du complément, et accéder à ces données lors de la prochaine ouverture du complément. Vous pouvez utiliser des techniques de programmation web courantes pour enregistrer les données, comme les cookies de navigateur ou le stockage web HTML 5. Si votre complément est également exécuté dans Excel, PowerPoint ou Word, vous pouvez également utiliser les méthodes de l’objet [Settings](/javascript/api/office/office.settings). Les données créées avec l’objet **Settings** sont stockées dans la feuille de calcul, la présentation ou le document dans lequel le complément a été inséré et enregistré. Ces données sont disponibles seulement pour le complément qui les a créées.

Pour éviter les allers-retours vers le serveur où le document est stocké, les données créées avec l'objet **Settings** sont gérées en mémoire au moment de l'exécution. Les données de paramètres enregistrées précédemment sont chargées en mémoire lors de l’initialisation du complément et les modifications apportées à ces données sont uniquement enregistrées dans le document quand vous appelez la méthode [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-). En interne, les données sont stockées dans un objet JSON sérialisé en tant que paires nom/valeur. Vous pouvez utiliser les méthodes [get](/javascript/api/office/office.settings#get-name-), [set](/javascript/api/office/office.settings#set-name--value-) et [remove](/javascript/api/office/office.settings#remove-name-) de l’objet **Settings** pour lire, écrire et supprimer des éléments dans la copie en mémoire des données. La ligne de code suivante explique comment créer un paramètre nommé `themeColor` et définir sa valeur sur « green ».




```js
Office.context.document.settings.set('themeColor', 'green');
```

Étant donné que les données de paramètres créées ou supprimées avec les méthodes **set** et **remove** agissent sur une copie en mémoire des données, vous devez appeler **saveAsync** pour rendre persistantes les modifications apportées aux données de paramètres dans le document utilisé par votre complément.

Pour plus de détails sur l’utilisation des données personnalisées à l’aide des méthodes de l’objet **Settings**, voir la page relative à la [conservation des données et aux paramètres d’état de complément](persisting-add-in-state-and-settings.md).


## <a name="reading-properties-of-a-project-document"></a>Lecture des propriétés d’un document de projet

Si votre complément de volet Office s’exécute dans Project, vous pouvez lire les données de certains champs, ressources et champs de tâche du projet actif. Pour ce faire, vous pouvez utiliser les méthodes et les événements de l’objet [ProjectDocument](/javascript/api/office/office.document), qui étend l’objet **Document** pour fournir des fonctionnalités supplémentaires propres à Project.

Pour des exemples de lecture de données Project, voir [Créer votre premier complément du volet Office pour Projet 2013 à l’aide d’un éditeur de texte](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).


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

- [API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office)
- [Référence de schéma pour les manifestes des compléments Office](../develop/add-in-manifests.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
