---
title: "Prise en charge de l’API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office\_2013"
description: Utilisez l Office API JavaScript pour créer un volet De tâches dans Office 2013.
ms.date: 07/08/2021
ms.localizationpriority: medium
---

# <a name="office-javascript-api-support-for-content-and-task-pane-add-ins-in-office-2013"></a>Prise en charge de l’API JavaScript pour Office pour les compléments de contenu et du volet Office dans Office 2013

[!include[information about the common API](../includes/alert-common-api-info.md)]

Vous pouvez utiliser [l’API JavaScript Office pour](../reference/javascript-api-for-office.md) créer des modules de contenu ou du volet Des tâches pour Office applications clientes 2013. Les méthodes et les objets pris en charge par les compléments du volet Office et de contenu sont classés de la manière suivante :

1. **Objets communs partagés avec d’autres Office des modules.** Ces objets incluent [Office](/javascript/api/office), [Context](/javascript/api/office/office.context) et [AsyncResult](/javascript/api/office/office.asyncresult). L’objet `Office` est l’objet racine de l Office API JavaScript. L’objet `Context` représente l’environnement d’runtime du add-in. Les `Office` deux objets `Context` sont fondamentaux pour n’importe quel Office de recherche. L’objet `AsyncResult` représente les résultats d’une opération asynchrone, `getSelectedDataAsync` telle que les données renvoyées à la méthode, qui lit ce qu’un utilisateur a sélectionné dans un document.

2. **Objet Document.** La majorité des éléments de l’API disponibles pour les compléments de contenu et du volet Office sont exposés via les méthodes, propriétés et événements de l’objet [Document](/javascript/api/office/office.document). Un application de contenu ou du volet Des tâches peut utiliser la propriété [Office.context.document](/javascript/api/office/office.context#office-office-context-document-member) pour accéder à l’objet **Document** et, par son biais, accéder aux membres clés de l’API pour utiliser des données dans des documents, tels que les objets [Bindings](/javascript/api/office/office.bindings) et [CustomXmlParts](/javascript/api/office/office.customxmlparts), et les méthodes [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)), [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) et [getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)). L’objet `Document` fournit également la propriété [mode](/javascript/api/office/office.document#office-office-document-mode-member) permettant de déterminer si un document est en lecture seule ou en mode édition, la propriété [d’URL](/javascript/api/office/office.document#office-office-document-url-member) pour obtenir l’URL du document actuel et l’accès à l’objet [Paramètres](/javascript/api/office/office.settings). L’objet `Document` prend également en charge l’ajout de handlers d’événements pour [l’événement SelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) , afin que vous pouvez détecter quand un utilisateur modifie sa sélection dans le document.

   Un add-in `Document` de contenu ou du volet Des tâches peut accéder à l’objet uniquement après le chargement du DOM et de l’environnement d’runtime, généralement dans le handler d’événements pour l’événement [Office.initialize](/javascript/api/office). Pour plus d’informations sur le flux d’événements lors de l’initialisation d’un complément et sur la vérification du chargement correct du DOM et de l’environnement d’exécution, voir la page relative au [chargement du DOM et de l’environnement d’exécution](loading-the-dom-and-runtime-environment.md).

3. **Objets pour l’utilisation de fonctionnalités spécifiques.** Pour utiliser des fonctionnalités spécifiques de l’API, utilisez les objets et méthodes suivants.

    - Les objets [CustomXmlParts](/javascript/api/office/office.bindings), [CustomXmlPart](/javascript/api/office/office.binding) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.

    - Les objets [CustomXmlParts](/javascript/api/office/office.customxmlparts) et [CustomXmlPart](/javascript/api/office/office.customxmlpart) et les objets associés pour créer et manipuler des parties XML personnalisées dans des documents Word.

    - Les objets [File](/javascript/api/office/office.file) et [Slice](/javascript/api/office/office.slice) pour créer une copie de l’intégralité du document, le diviser en blocs ou en « sections », puis lire ou transmettre les données dans ces sections.

    - L’objet [Settings](/javascript/api/office/office.settings) pour enregistrer des données personnalisées, telles que des préférences utilisateur et l’état du complément.

> [!IMPORTANT]
> Certains des membres d’API ne sont pas pris en charge dans toutes les applications Office pouvant héberger des compléments de contenu et du volet Office. Pour déterminer les membres pris en charge, voir les ressources suivantes :

Pour obtenir un résumé de Office prise en charge de l’API JavaScript dans Office applications clientes, voir [Understanding the Office JavaScript API](understanding-the-javascript-api-for-office.md).

## <a name="read-and-write-to-an-active-selection-in-a-document-spreadsheet-or-presentation"></a>Lire et écrire dans une sélection active dans un document, une feuille de calcul ou une présentation

Vous pouvez lire ou écrire dans la sélection en cours de l’utilisateur dans un document, une feuille de calcul ou une présentation. Selon l’application Office de votre application, vous pouvez spécifier le type de structure de données à lire ou à écrire en tant que paramètre dans les méthodes [getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) et [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) de l’objet [Document](/javascript/api/office/office.document). Par exemple, vous pouvez indiquer n’importe quel type de données (HTML, données tabulaires, Office Open XML ou texte) pour Word, des données texte et tabulaires pour Excel et des données texte pour PowerPoint et Project. Vous pouvez également créer des gestionnaires d’événements pour détecter les modifications apportées à la sélection de l’utilisateur. L’exemple suivant obtient les données de la sélection en tant que texte à l’aide de la `getSelectedDataAsync` méthode.


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

Vous pouvez utiliser les méthodes `getSelectedDataAsync` et les `setSelectedDataAsync` méthodes pour lire ou écrire dans la sélection actuelle  de l’utilisateur dans un document, une feuille de calcul ou une présentation. Toutefois, si vous souhaitez accéder à la même région dans un document via des sessions d’exécution de votre complément sans demander à l’utilisateur d’effectuer une sélection, vous devez d’abord établir une liaison avec cette région. Avec une liaison, vous pouvez également vous abonner à des données et à des événements de modification de sélection, uniquement pour la région liée.

Vous pouvez ajouter une liaison à l’aide des méthodes [addFromNamedItemAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromnameditemasync-member(1)), [addFromPromptAsync](/javascript/api/office/office.bindings#office-office-bindings-addfrompromptasync-member(1)) ou [addFromSelectionAsync](/javascript/api/office/office.bindings#office-office-bindings-addfromselectionasync-member(1)) de l’objet [Bindings](/javascript/api/office/office.bindings).

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

Si votre complément du volet Office s’exécute dans PowerPoint ou Word, vous pouvez utiliser les méthodes [Document.getFileAsync](/javascript/api/office/office.document#office-office-document-getfileasync-member(1)), [File.getSliceAsync](/javascript/api/office/office.file#office-office-file-getsliceasync-member(1)) et [File.closeAsync](/javascript/api/office/office.file#office-office-file-closeasync-member(1)) pour obtenir la totalité d’une présentation ou d’un document.

Lorsque vous appelez `Document.getFileAsync` , vous obtenez une copie du document dans un [objet](/javascript/api/office/office.file) File. L’objet `File` permet d’accéder au document en « blocs » représentés en tant [qu’objets Slice](/javascript/api/office/office.slice) . Lorsque vous `getFileAsync`appelez , vous pouvez spécifier le type de fichier (texte ou format XML Open Office compressé) et la taille des tranches (jusqu’à 4 Mo). Pour accéder au contenu de l’objet `File` , `File.getSliceAsync` vous appelez ensuite qui renvoie les données brutes dans la [propriété Slice.data](/javascript/api/office/office.slice#office-office-slice-data-member) . Si vous avez spécifié un format compressé, vous obtiendrez les données du fichier sous la forme d’un tableau d’octets. Si vous transférez le fichier à un service web, vous pouvez transformer les données brutes compressées dans une chaîne codée en Base64 avant l’envoi. Enfin, lorsque vous avez terminé d’obtenir des tranches du fichier, utilisez la `File.closeAsync` méthode pour fermer le document.

Pour plus d’informations, reportez-vous à l’article expliquant comment [obtenir l’intégralité d’un document à partir d’un complément pour PowerPoint ou Word](../word/get-the-whole-document-from-an-add-in-for-word.md).

## <a name="read-and-write-custom-xml-parts-of-a-word-document"></a>Lire et écrire des parties XML personnalisées d’un document Word

Grâce aux contrôles de contenu et au format de fichier Office Open XML, vous pouvez ajouter des parties XML personnalisées à un document Word et lier des éléments dans les parties XML aux contrôles de contenu de ce document. Lorsque vous ouvrez le document, Word lit et remplit automatiquement les contrôles de contenu liés avec les données des parties XML personnalisées. Les utilisateurs peuvent également écrire des données dans les contrôles de contenu. Lorsqu’ils enregistrent le document, les données des contrôles sont alors enregistrées dans les parties XML liées. Si votre complément du volet Office s’exécute dans Word, vous pouvez utiliser la propriété [Document.customXmlParts](/javascript/api/office/office.document#office-office-document-customxmlparts-member), ainsi que les objets [CustomXmlParts](/javascript/api/office/office.customxmlparts), [CustomXmlPart](/javascript/api/office/office.customxmlpart) et [CustomXmlNode](/javascript/api/office/office.customxmlnode) pour lire et écrire des données de manière dynamique dans le document.

Les parties XML personnalisées peuvent être associées à des espaces de noms. Pour obtenir des données à partir des parties XML personnalisées dans un espace de noms, utilisez la méthode [CustomXmlParts.getByNamespaceAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbynamespaceasync-member(1)).

Vous pouvez également utiliser la [CustomXmlParts.getByIdAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-getbyidasync-member(1)) pour accéder aux parties XML personnalisées par leur GUID. Après avoir obtenu une partie XML personnalisée, utilisez la méthode [CustomXmlPart.getXmlAsync](/javascript/api/office/office.customxmlpart#office-office-customxmlpart-getxmlasync-member(1)) pour obtenir les données XML.

Pour ajouter une nouvelle partie XML personnalisée à un document, `Document.customXmlParts` utilisez la propriété pour obtenir les parties XML personnalisées du document et appelez la méthode [CustomXmlParts.addAsync](/javascript/api/office/office.customxmlparts#office-office-customxmlparts-addasync-member(1)) .

Pour obtenir des informations détaillées sur l’utilisation de parties XML personnalisées avec un complément du volet Office, voir [Création de meilleurs compléments pour Word avec Office Open XML](../word/create-better-add-ins-for-word-with-office-open-xml.md).

## <a name="persisting-add-in-settings"></a>Persistance des paramètres de complément

Vous devez souvent enregistrer les données personnalisées pour votre complément, telles que les préférences d’un utilisateur ou l’état du complément, et accéder à ces données lors de la prochaine ouverture du complément. Vous pouvez utiliser des techniques de programmation web courantes pour enregistrer les données, comme les cookies de navigateur ou le stockage web HTML 5. Si votre complément est également exécuté dans Excel, PowerPoint ou Word, vous pouvez également utiliser les méthodes de l’objet [Settings](/javascript/api/office/office.settings). Les données créées avec l’objet `Settings` sont stockées dans la feuille de calcul, la présentation ou le document avec qui le module a été inséré et enregistré. Ces données sont disponibles seulement pour le complément qui les a créées.

Pour éviter les allers-retours vers le serveur où le document est stocké, `Settings` les données créées avec l’objet sont gérées en mémoire au moment de l’exécuter. Les données de paramètres enregistrées précédemment sont chargées en mémoire lors de l’initialisation du complément et les modifications apportées à ces données sont uniquement enregistrées dans le document quand vous appelez la méthode [Settings.saveAsync](/javascript/api/office/office.settings#office-office-settings-saveasync-member(1)). En interne, les données sont stockées dans un objet JSON sérialisé en tant que paires nom/valeur. Vous pouvez utiliser les méthodes [get](/javascript/api/office/office.settings#office-office-settings-get-member(1)), [set](/javascript/api/office/office.settings#office-office-settings-set-member(1)) et [remove](/javascript/api/office/office.settings#office-office-settings-remove-member(1)) de l’objet **Settings** pour lire, écrire et supprimer des éléments dans la copie en mémoire des données. La ligne de code suivante explique comment créer un paramètre nommé `themeColor` et définir sa valeur sur « green ».

```js
Office.context.document.settings.set('themeColor', 'green');
```

`set` `remove` Étant donné que les données de paramètres créées ou supprimées avec les méthodes agissent sur une copie en mémoire des données, `saveAsync` vous devez appeler pour faire persister les modifications apportées aux données de paramètres dans le document sur le document avec qui votre module est en cours d’utilisation.

Pour plus d’informations sur l’utilisation de données personnalisées `Settings` à l’aide des méthodes de l’objet, voir [Persisting add-in state and settings](persisting-add-in-state-and-settings.md).

## <a name="read-properties-of-a-project-document"></a>Lire les propriétés d’un document de projet

Si votre complément de volet Office s’exécute dans Project, vous pouvez lire les données de certains champs, ressources et champs de tâche du projet actif. Pour ce faire, vous utilisez les méthodes et les événements de l’objet [ProjectDocument](/javascript/api/office/office.document), `Document` qui étend l’objet pour fournir des fonctionnalités Project spécifiques.

Pour des exemples de lecture de données Project, voir [Créer votre premier complément du volet Office pour Projet 2013 à l’aide d’un éditeur de texte](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).

## <a name="permissions-model-and-governance"></a>Modèle d’autorisations et gouvernance

Votre add-in utilise l’élément `Permissions` dans son manifeste pour demander l’autorisation d’accéder au niveau de fonctionnalité dont il a besoin à partir de l Office API JavaScript. Par exemple, si votre add-in nécessite un accès en lecture/écriture au document, `ReadWriteDocument` son manifeste doit spécifier comme valeur de texte dans son `Permissions` élément. Étant donné que les autorisations ont pour objectif de protéger la vie privée et la sécurité de l’utilisateur, en tant que meilleures pratiques, nous vous recommandons de demander le niveau d’autorisation minimal requis pour ses fonctionnalités. L’exemple suivant illustre la demande de l’autorisation **ReadDocument** dans le manifeste d’un complément du volet Office.

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

Pour plus d’informations, voir [Demande d’autorisations pour l’utilisation d’API dans les modules complémentaires](requesting-permissions-for-api-use-in-content-and-task-pane-add-ins.md).

## <a name="see-also"></a>Voir aussi

- [API JavaScript pour Office](../reference/javascript-api-for-office.md)
- [Référence de schéma pour les manifestes des compléments Office](../develop/add-in-manifests.md)
- [Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office](../testing/testing-and-troubleshooting.md)
