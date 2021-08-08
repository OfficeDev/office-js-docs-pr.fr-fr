---
title: Prise en charge de l’API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013
description: Utilisez l Office API JavaScript pour créer un volet De tâches dans Office 2013.
ms.date: 07/08/2021
localization_priority: Normal
ms.openlocfilehash: 644bc1f0759d381de412cb276a1535d2251abb0a6a0be78b45d9cc0a245758c7
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2021
ms.locfileid: "57079994"
---
# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Prise en charge de l’API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

Vous pouvez utiliser [l’API JavaScript Office pour](../reference/javascript-api-for-office.md) créer des modules de contenu ou du volet Des tâches pour Office applications clientes 2013. Les méthodes et les objets pris en charge par les compléments du volet Office et de contenu sont classés de la manière suivante :

1. **Objets communs partagés avec d’autres Office des modules.** Ces objets incluent [Office,](/javascript/api/office) [Context](/javascript/api/office/office.context)et [AsyncResult](/javascript/api/office/office.asyncresult). `Office`L’objet est l’objet racine de l Office API JavaScript. `Context`L’objet représente l’environnement d’runtime du add-in. Les deux objets sont fondamentaux pour n’importe `Office` `Context` quel Office de recherche. L’objet représente les résultats d’une opération asynchrone, telle que les données renvoyées à la méthode, qui lit ce qu’un utilisateur a sélectionné `AsyncResult` `getSelectedDataAsync` dans un document.

2. **Objet Document.** La majorité des éléments de l’API disponibles pour les compléments de contenu et du volet Office sont exposés via les méthodes, propriétés et événements de l’objet [Document](/javascript/api/office/office.document). Un application de contenu ou du volet Des tâches peut utiliser la propriété [Office.context.document](/javascript/api/office/office.context#document) pour accéder à l’objet **Document** et, par son biais, accéder aux membres clés de l’API pour utiliser des données dans des documents, tels que les objets [Bindings](/javascript/api/office/office.bindings) et [CustomXmlParts,](/javascript/api/office/office.customxmlparts) et les méthodes [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_), [setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_)et [getFileAsync.](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_) L’objet fournit également la propriété mode permettant de déterminer si un document est en lecture seule ou en mode édition, la propriété d’URL permettant d’obtenir l’URL du document actuel et l’accès à `Document` l’objet [Paramètres.](/javascript/api/office/office.settings) [](/javascript/api/office/office.document#mode) [](/javascript/api/office/office.document#url) L’objet prend également en charge l’ajout de handlers d’événements pour `Document` l’événement [SelectionChanged,](/javascript/api/office/office.documentselectionchangedeventargs) afin que vous pouvez détecter quand un utilisateur modifie sa sélection dans le document.

   Un add-in de contenu ou du volet Des tâches peut accéder à l’objet uniquement après le chargement du DOM et de l’environnement d’runtime, généralement dans le handler d’événements pour l’événement `Document` [Office.initialize.](/javascript/api/office) Pour plus d’informations sur le flux d’événements lors de l’initialisation d’un complément et sur la vérification du chargement correct du DOM et de l’environnement d’exécution, voir la page relative au [chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

3. **Objets pour l’utilisation de fonctionnalités spécifiques.** Pour utiliser des fonctionnalités spécifiques de l’API, utilisez les objets et méthodes suivants.

    - Les objets [CustomXmlParts](/javascript/api/office/office.bindings), [CustomXmlPart](/javascript/api/office/office.binding) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.

    - Les objets [CustomXmlParts](/javascript/api/office/office.customxmlparts) et [CustomXmlPart](/javascript/api/office/office.customxmlpart) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.

    - Les objets [File](/javascript/api/office/office.file) et [Slice](/javascript/api/office/office.slice) pour créer une copie de l’intégralité du document, le diviser en blocs ou en « sections », puis lire ou transmettre les données dans ces sections.

    - L’objet [Settings](/javascript/api/office/office.settings) pour enregistrer des données personnalisées, telles que des préférences utilisateur et l’état du complément.

> [!IMPORTANT]
> Certains des membres d’API ne sont pas pris en charge dans toutes les applications Office pouvant héberger des compléments de contenu et du volet Office. Pour déterminer les membres pris en charge, voir les ressources suivantes :

Pour obtenir un résumé de Office prise en charge de l’API JavaScript dans Office applications clientes, voir [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md).

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>Lire et écrire dans une sélection active dans un document, une feuille de calcul ou une présentation

Vous pouvez lire ou écrire dans la sélection en cours de l’utilisateur dans un document, une feuille de calcul ou une présentation. Selon l’application Office de votre application, vous pouvez spécifier le type de structure de données à lire ou à écrire en tant que paramètre dans les méthodes [getSelectedDataAsync](/javascript/api/office/office.document#getSelectedDataAsync_coercionType__options__callback_) et [setSelectedDataAsync](/javascript/api/office/office.document#setSelectedDataAsync_data__options__callback_) de l’objet [Document.](/javascript/api/office/office.document) Par exemple, vous pouvez indiquer n’importe quel type de données (HTML, données tabulaires, Office Open XML ou texte) pour Word, des données texte et tabulaires pour Excel et des données texte pour PowerPoint et Project. Vous pouvez également créer des gestionnaires d’événements pour détecter les modifications apportées à la sélection de l’utilisateur. L’exemple suivant obtient les données de la sélection en tant que texte à l’aide de `getSelectedDataAsync` la méthode.


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

## <a name="bind-to-a-region-in-a-document-or-spreadsheet"></a>Liaison à une région dans un document ou une feuille de calcul

Vous pouvez utiliser les méthodes et les méthodes pour lire ou écrire dans la sélection actuelle de l’utilisateur dans un document, une feuille de calcul `getSelectedDataAsync` `setSelectedDataAsync` ou une présentation.  Toutefois, si vous souhaitez accéder à la même région dans un document via des sessions d’exécution de votre complément sans demander à l’utilisateur d’effectuer une sélection, vous devez d’abord établir une liaison avec cette région. Avec une liaison, vous pouvez également vous abonner à des données et à des événements de modification de sélection, uniquement pour la région liée.

Vous pouvez ajouter une liaison à l’aide des méthodes [addFromNamedItemAsync](/javascript/api/office/office.bindings#addFromNamedItemAsync_itemName__bindingType__options__callback_), [addFromPromptAsync](/javascript/api/office/office.bindings#addFromPromptAsync_bindingType__options__callback_) ou [addFromSelectionAsync](/javascript/api/office/office.bindings#addFromSelectionAsync_bindingType__options__callback_) de l’objet [Bindings](/javascript/api/office/office.bindings).

Voici un exemple qui ajoute une liaison au texte actuellement sélectionné dans un document, à l’aide de la `Bindings.addFromSelectionAsync` méthode.

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

## <a name="get-entire-documents"></a>Obtenir des documents entiers

Si votre complément du volet Office s’exécute dans PowerPoint ou Word, vous pouvez utiliser les méthodes [Document.getFileAsync](/javascript/api/office/office.document#getFileAsync_fileType__options__callback_), [File.getSliceAsync](/javascript/api/office/office.file#getSliceAsync_sliceIndex__callback_) et [File.closeAsync](/javascript/api/office/office.file#closeAsync_callback_) pour obtenir la totalité d’une présentation ou d’un document.

Lorsque vous `Document.getFileAsync` appelez, vous obtenez une copie du document dans un [objet](/javascript/api/office/office.file) File. `File`L’objet permet d’accéder au document en « blocs » représentés en tant [qu’objets Slice.](/javascript/api/office/office.slice) Lorsque vous appelez , vous pouvez spécifier le type de fichier (texte ou format XML Open Office compressé) et la taille des `getFileAsync` tranches (jusqu’à 4 Mo). Pour accéder au contenu de l’objet, vous appelez ensuite qui renvoie les données brutes `File` `File.getSliceAsync` dans la propriété [Slice.data.](/javascript/api/office/office.slice#data) Si vous avez spécifié un format compressé, vous obtiendrez les données du fichier sous la forme d’un tableau d’octets. Si vous transférez le fichier à un service web, vous pouvez transformer les données brutes compressées dans une chaîne codée en Base64 avant l’envoi. Enfin, lorsque vous avez terminé d’obtenir des tranches du fichier, utilisez `File.closeAsync` la méthode pour fermer le document.

Pour plus d’informations, reportez-vous à l’article expliquant comment [obtenir l’intégralité d’un document à partir d’un complément pour PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>Lire et écrire des parties XML personnalisées d’un document Word

Grâce aux contrôles de contenu et au format de fichier Office Open XML, vous pouvez ajouter des parties XML personnalisées à un document Word et lier des éléments dans les parties XML aux contrôles de contenu de ce document. Lorsque vous ouvrez le document, Word lit et remplit automatiquement les contrôles de contenu liés avec les données des parties XML personnalisées. Les utilisateurs peuvent également écrire des données dans les contrôles de contenu. Lorsqu’ils enregistrent le document, les données des contrôles sont alors enregistrées dans les parties XML liées. Si votre complément du volet Office s’exécute dans Word, vous pouvez utiliser la propriété [Document.customXmlParts](/javascript/api/office/office.document#customXmlParts), ainsi que les objets [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) et [CustomXmlNode](/javascript/api/office/office.customxmlnode) pour lire et écrire des données de manière dynamique dans le document.

Les parties XML personnalisées peuvent être associées à des espaces de noms. Pour obtenir des données à partir des parties XML personnalisées dans un espace de noms, utilisez la méthode [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#getByNamespaceAsync_ns__options__callback_).

Vous pouvez également utiliser la [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#getByIdAsync_id__options__callback_) pour accéder aux parties XML personnalisées par leur GUID. Après avoir obtenu une partie XML personnalisée, utilisez la méthode [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#getXmlAsync_options__callback_) pour obtenir les données XML.

Pour ajouter une nouvelle partie XML personnalisée à un document, utilisez la propriété pour obtenir les parties XML personnalisées du document et appelez la méthode `Document.customXmlParts` [CustomXmlParts.addAsync.](/javascript/api/office/office.customxmlparts#addAsync_xml__options__callback_)

Pour obtenir des informations détaillées sur l’utilisation de parties XML personnalisées avec un complément du volet Office, voir [Création de meilleurs compléments pour Word avec Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).

## <a name="persisting-add-in-settings"></a>Persistance des paramètres de complément

Vous devez souvent enregistrer les données personnalisées pour votre complément, telles que les préférences d’un utilisateur ou l’état du complément, et accéder à ces données lors de la prochaine ouverture du complément. Vous pouvez utiliser des techniques de programmation web courantes pour enregistrer les données, comme les cookies de navigateur ou le stockage web HTML 5. Si votre complément est également exécuté dans Excel, PowerPoint ou Word, vous pouvez également utiliser les méthodes de l’objet [Settings](/javascript/api/office/office.settings). Les données créées avec l’objet sont stockées dans la feuille de calcul, la présentation ou le document avec qui le module a été inséré et `Settings` enregistré. Ces données sont disponibles seulement pour le complément qui les a créées.

Pour éviter les allers-retours vers le serveur où le document est stocké, les données créées avec l’objet sont gérées en mémoire au moment `Settings` de l’exécuter. Les données de paramètres enregistrées précédemment sont chargées en mémoire lors de l’initialisation du complément et les modifications apportées à ces données sont uniquement enregistrées dans le document quand vous appelez la méthode [Settings.saveAsync](/javascript/api/office/office.settings#saveAsync_options__callback_). En interne, les données sont stockées dans un objet JSON sérialisé en tant que paires nom/valeur. Vous pouvez utiliser les méthodes [get](/javascript/api/office/office.settings#get_name_), [set](/javascript/api/office/office.settings#set_name__value_) et [remove](/javascript/api/office/office.settings#remove_name_) de l’objet **Settings** pour lire, écrire et supprimer des éléments dans la copie en mémoire des données. La ligne de code suivante explique comment créer un paramètre nommé `themeColor` et définir sa valeur sur « green ».

```js
Office.context.document.settings.set('themeColor', 'green');
```

Étant donné que les données de paramètres créées ou supprimées avec les méthodes agissent sur une copie en mémoire des données, vous devez appeler pour faire persister les modifications apportées aux données de paramètres dans le document sur le document avec qui votre module est en cours `set` `remove` `saveAsync` d’utilisation.

Pour plus d’informations sur l’utilisation de données personnalisées à l’aide des méthodes de l’objet, voir Persistance de l’état et `Settings` [des paramètres du module.](persisting-add-in-state-and-settings.md)

## <a name="read-properties-of-a-project-document"></a>Lire les propriétés d’un document de projet

Si votre complément de volet Office s’exécute dans Project, vous pouvez lire les données de certains champs, ressources et champs de tâche du projet actif. Pour ce faire, vous utilisez les méthodes et les événements de l’objet [ProjectDocument,](/javascript/api/office/office.document) qui étend l’objet pour fournir des fonctionnalités Project `Document` spécifiques.

Pour des exemples de lecture de données Project, voir [Créer votre premier complément du volet Office pour Projet 2013 à l’aide d’un éditeur de texte](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).

## <a name="permissions-model-and-governance"></a>Modèle d’autorisations et gouvernance

Votre add-in utilise l’élément dans son manifeste pour demander l’autorisation d’accéder au niveau de fonctionnalité dont il a besoin à partir de l Office `Permissions` API JavaScript. Par exemple, si votre add-in nécessite un accès en lecture/écriture au document, son manifeste doit spécifier comme valeur de texte `ReadWriteDocument` dans son `Permissions` élément. Étant donné que les autorisations ont pour objectif de protéger la vie privée et la sécurité de l’utilisateur, en tant que meilleures pratiques, nous vous recommandons de demander le niveau d’autorisation minimal requis pour ses fonctionnalités. L’exemple suivant illustre la demande de l’autorisation **ReadDocument** dans le manifeste d’un complément du volet Office.

```XML
<?xml version="1.0" encoding="utf-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
 xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" 
 xsi:type="TaskPaneApp">
???<!-- Other manifest elements omitted. -->
  <Permissions>ReadDocument</Permissions>
???
</OfficeApp>

```

Pour plus d’informations, voir [Demande d’autorisations pour l’utilisation d’API dans les modules complémentaires.](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md)

## <a name="see-also"></a>Voir aussi

- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
- [Référence de schéma pour les manifestes des compléments Office](../develop/add-in-manifests.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
