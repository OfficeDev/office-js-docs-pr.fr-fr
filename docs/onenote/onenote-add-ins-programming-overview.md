---
title: Vue d’ensemble de la programmation de l’API JavaScript de OneNote
description: En savoir plus sur l’API JavaScript de OneNote pour les compléments OneNote sur le web.
ms.date: 07/18/2022
ms.topic: overview
ms.custom: scenarios:getting-started
ms.localizationpriority: high
ms.openlocfilehash: d44a01cf0f676057ca072cff74e2e80057f645f4
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092909"
---
# <a name="onenote-javascript-api-programming-overview"></a>Vue d’ensemble de la programmation de l’API JavaScript de OneNote

OneNote introduces a JavaScript API for OneNote add-ins on the web. You can create task pane add-ins, content add-ins, and add-in commands that interact with OneNote objects and connect to web services or other web-based resources.

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="components-of-an-office-add-in"></a>Composants d’un complément Office

Les compléments sont constitués de deux composants de base :

- A **web application** consisting of a webpage and any required JavaScript, CSS, or other files. These files are hosted on a web server or web hosting service, such as Microsoft Azure. In OneNote on the web, the web application displays in a browser control or iframe.

- An **XML manifest** that specifies the URL of the add-in's webpage and any access requirements, settings, and capabilities for the add-in. This file is stored on the client. OneNote add-ins use the same [manifest](../develop/add-in-manifests.md) format as other Office Add-ins.

### <a name="office-add-in--manifest--webpage"></a>Complément pour Office = manifeste + page web

![Le complément Office se compose d’un manifeste et d’une page web.](../images/onenote-add-in.png)

## <a name="using-the-javascript-api"></a>Utilisation de l’API JavaScript

Add-ins use the runtime context of the Office application to access the JavaScript API. The API has two layers:

- Une **API spécifique à l’application** pour les opérations spécifiques de OneNote, accessible via l’objet`Application`Application.
- Une **API commune** qui est partagée entre les applications Office, accessible via l’objet `Document`.

### <a name="accessing-the-application-specific-api-through-the-application-object"></a>Accès à l’API spécifique à l’application via l’objet *Application*

Use the `Application` object to access OneNote objects such as **Notebook**, **Section**, and **Page**. With application-specific APIs, you run batch operations on proxy objects. The basic flow goes something like this:

1. Obtenir l’instance de l’application à partir du contexte.

2. Create a proxy that represents the OneNote object you want to work with. You interact synchronously with proxy objects by reading and writing their properties and calling their methods.

3. Call `load` on the proxy to fill it with the property values specified in the parameter. This call is added to the queue of commands.

   > [!NOTE]
   > Les appels de méthode à l’API (tels que `context.application.getActiveSection().pages;`) sont également ajoutés à la file d’attente.

4. Call `context.sync` to run all queued commands in the order that they were queued. This synchronizes the state between your running script and the real objects, and by retrieving properties of loaded OneNote objects for use in your script. You can use the returned promise object for chaining additional actions.

Par exemple :

```js
async function getPagesInSection() {
    await OneNote.run(async (context) => {

        // Get the pages in the current section.
        const pages = context.application.getActiveSection().pages;

        // Queue a command to load the id and title for each page.
        pages.load('id,title');

        // Run the queued commands, and return a promise to indicate task completion.
        await context.sync();
            
        // Read the id and title of each page.
        $.each(pages.items, function(index, page) {
            let pageId = page.id;
            let pageTitle = page.title;
            console.log(pageTitle + ': ' + pageId);
        });
    });
}
```

Pour en savoir plus sur le modèle `load`/`sync` et d’autres pratiques courantes dans les API JavaScript OneNote, consultez [l’utilisation du modèle API spécifique à l’application](../develop/application-specific-api-model.md).

Vous pouvez déterminer les objets et les opérations OneNote pris en charge dans la [référence de l’API](../reference/overview/onenote-add-ins-javascript-reference.md).

#### <a name="onenote-javascript-api-requirement-sets"></a>Ensembles de conditions requises de l’API JavaScript pour OneNote

Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For detailed information about OneNote JavaScript API requirement sets, see [OneNote JavaScript API requirement sets](/javascript/api/requirement-sets/onenote/onenote-api-requirement-sets).

### <a name="accessing-the-common-api-through-the-document-object"></a>Accès à l’API commune via l’objet *Document*

Utilisez l’objet `Document` pour accéder à l’API commune, par exemple les méthodes[getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) et [setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)).

Par exemple :  

```js
function getSelectionFromPage() {
    Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        { valueFormat: "unformatted" },
        function (asyncResult) {
            const error = asyncResult.error;
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                console.log(error.message);
            }
            else $('#input').val(asyncResult.value);
        });
}
```

Les compléments OneNote prennent en charge uniquement les API communes suivantes.

| API | Commentaires |
|:------|:------|
| [Office.context.document.getSelectedDataAsync](/javascript/api/office/office.document#office-office-document-getselecteddataasync-member(1)) | Office.CoercionType.Text`Office.CoercionType.Text` et Office.CoercionType.Matrix`Office.CoercionType.Matrix` uniquement |
| [Office.context.document.setSelectedDataAsync](/javascript/api/office/office.document#office-office-document-setselecteddataasync-member(1)) | `Office.CoercionType.Text`, `Office.CoercionType.Image`et `Office.CoercionType.Html` uniquement |
| [var mySetting = Office.context.document.settings.get(name);](/javascript/api/office/office.settings#office-office-settings-get-member(1)) | Les paramètres sont pris en charge par les compléments de contenu uniquement |
| [Office.context.document.settings.set(name, value);](/javascript/api/office/office.settings#office-office-settings-set-member(1)) | Les paramètres sont pris en charge par les compléments de contenu uniquement |
| [Office.EventType.DocumentSelectionChanged](/javascript/api/office/office.documentselectionchangedeventargs) |*Aucun.*|

En règle générale, vous utilisez l’API commune pour effectuer une action qui n’est pas prise en charge dans l’API spécifique à l’application. Pour plus d’informations sur les API communes, voir le [Modèle d’objet API JavaScript communes](../develop/office-javascript-api-object-model.md).

<a name="om-diagram"></a>

## <a name="onenote-object-model-diagram"></a>Diagramme du modèle objet OneNote

Le diagramme suivant représente ce qui est actuellement disponible dans l’API JavaScript de OneNote.

  ![Diagramme de modèle objet OneNote.](../images/onenote-om.png)

## <a name="see-also"></a>Voir aussi

- [Développement de compléments Office](../develop/develop-overview.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Créer votre premier complément OneNote](../quickstarts/onenote-quickstart.md)
- [Référence de l’API JavaScript de OneNote](../reference/overview/onenote-add-ins-javascript-reference.md)
- [Exemple de grille d’évaluation](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Vue d’ensemble de la plateforme des compléments Office](../overview/office-add-ins.md)
