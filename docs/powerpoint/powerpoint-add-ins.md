---
title: Compléments PowerPoint
description: Découvrez comment utiliser des compléments PowerPoint afin de créer des solutions attrayantes pour les présentations sur différentes plateformes, notamment Windows, iPad et Mac, ainsi que dans un navigateur.
ms.date: 10/14/2020
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 476f8f34bc47d85842d5b31f8a0298bf2d5d7b18
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740839"
---
# <a name="powerpoint-add-ins"></a>Compléments PowerPoint

Vous pouvez utiliser des compléments PowerPoint afin de créer des solutions attrayantes pour les présentations de vos utilisateurs sur différentes plateformes, notamment Windows, iPad et Mac, ainsi que dans un navigateur. Vous pouvez créer deux types de commandes de complément PowerPoint:

- Utilisez des **compléments de contenu** pour ajouter du contenu HTML5 dynamique à vos présentations. Par exemple, consultez le complément [Diagrammes LucidChart pour PowerPoint](https://appsource.microsoft.com/product/office/wa104380117), qui vous permet d’injecter un diagramme interactif de LucidChart dans votre support de présentation.

- Utilisez des **compléments de volet Office** pour faire apparaître des informations de référence ou insérer des données dans la diapositive via un service. Par exemple, consultez le complément [Stock Photos gratuit Pexels](https://appsource.microsoft.com/product/office/wa104379997), qui vous permet d’ajouter des photos professionnelles à votre présentation.

## <a name="powerpoint-add-in-scenarios"></a>Scénarios de complément PowerPoint

Les exemples de code figurant dans l’article vous présentent certaines tâches de base en matière de développement de compléments de contenu pour PowerPoint. Notez également ce qui suit:

- Pour afficher des informations, ces exemples dépendent de la fonction`app.showNotification`, qui est incluse dans les modèles de projet de compléments Office Visual Studio. Si vous n’utilisez pas Visual Studio pour développer votre complément, vous devrez remplacer la fonction`showNotification`par votre propre code.

- Plusieurs de ces exemples dépendent également de l’objet`Globals` qui est déclaré en dehors de la portée de ces fonctions: `var Globals = {activeViewHandler:0, firstSlideId:0};`

- Pour utiliser ces exemples, votre projet complément doit [référencer Office.js version 1.1 bibliothèque ou version ultérieure](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Détecter l’affichage actif de la présentation et gérer l’événement ActiveViewChanged

Si vous créez un complément de contenu, vous devrez obtenir la vue active de la présentation et gérer`ActiveViewChanged`l’événement ActiveViewChanged dans le cadre de votre`Office.Initialize`gestionnaire.

> [!NOTE]
> Dans PowerPoint sur le web, l’événement [Document.ActiveViewChanged](/javascript/api/office/office.document) ne se déclenche jamais, car le mode Diaporama est considéré comme une nouvelle session. Dans ce cas, le complément doit extraire la vue active lors du chargement, comme indiqué ci-dessous.

Collez le code suivant:

- La fonction`getActiveFileView` appelle la méthode [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) afin de renvoyer si la vue actuelle de la présentation est une vue de « modification » (toutes les vues dans lesquelles vous modifiez des diapositives, telles que les vues **Normal** ou **Mode Plan**) ou « lecture » ( **Diaporama**ou**Mode Lecture**).

- La fonction`registerActiveViewChanged`appelle la méthode [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) afin d’inscrire un gestionnaire pour l’événement [Document.ActiveViewChanged](/javascript/api/office/office.document).


```js
//general Office.initialize function. Fires on load of the add-in.
Office.initialize = function(){

    //Gets whether the current view is edit or read.
    var currentView = getActiveFileView();

    //register for the active view changed handler
    registerActiveViewChanged();

    //render the content based off of the currentView
    //....
}

function getActiveFileView()
{
    Office.context.document.getActiveViewAsync(function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification(asyncResult.value);
        }
    });

}

function registerActiveViewChanged() {
    Globals.activeViewHandler = function (args) {
        app.showNotification(JSON.stringify(args));
    }

    Office.context.document.addHandlerAsync(Office.EventType.ActiveViewChanged, Globals.activeViewHandler,
        function (asyncResult) {
            if (asyncResult.status == "failed") {
                app.showNotification("Action failed with error: " + asyncResult.error.message);
            }
            else {
                app.showNotification(asyncResult.status);
            }
        });
}
```

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a>Accéder à une diapositive spécifique dans la présentation

Dans l’exemple de code suivant, la fonction`getSelectedRange` appelle la méthode[Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) pour obtenir l’objet JSON renvoyé par`asyncResult.value`, qui contient un tableau nommé `slides`. La matrice`slides`contient les IDs, les titres et les indexes de plage sélectionnées de diapositives (ou de la diapositive active si plusieurs diapositives ne sont pas sélectionnées). Elle enregistre également l’id de la première diapositive dans la plage sélectionnée à une variable globale.

```js
function getSelectedRange() {
    // Get the id, title, and index of the current slide (or selected slides) and store the first slide id */
    Globals.firstSlideId = 0;

    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            Globals.firstSlideId = asyncResult.value.slides[0].id;
            app.showNotification(JSON.stringify(asyncResult.value));
        }
    });
}
```

Dans l’exemple de code suivant la fonction`goToFirstSlide`appelle la méthode[Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-)pour accéder à la première diapositive qui a été identifiée par la fonction`getSelectedRange`illustrée précédemment.

```js
function goToFirstSlide() {
    Office.context.document.goToByIdAsync(Globals.firstSlideId, Office.GoToType.Slide, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="navigate-between-slides-in-the-presentation"></a>Naviguer entre les diapositives de la présentation

Dans l’exemple de code suivant, la fonction`goToSlideByIndex` appelle la méthode `Document.goToByIdAsync` pour passer à la diapositive suivante dans la présentation.

```js
function goToSlideByIndex() {
    var goToFirst = Office.Index.First;
    var goToLast = Office.Index.Last;
    var goToPrevious = Office.Index.Previous;
    var goToNext = Office.Index.Next;

    Office.context.document.goToByIdAsync(goToNext, Office.GoToType.Index, function (asyncResult) {
        if (asyncResult.status == "failed") {
            app.showNotification("Action failed with error: " + asyncResult.error.message);
        }
        else {
            app.showNotification("Navigation successful");
        }
    });
}
```

## <a name="get-the-url-of-the-presentation"></a>Obtenir l’URL de la présentation

La fonction`getFileUrl` appelle la méthode [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) pour obtenir l’URL du fichier de présentation.

```js
function getFileUrl() {
    //Get the URL of the current file.
    Office.context.document.getFilePropertiesAsync(function (asyncResult) {
        var fileUrl = asyncResult.value.url;
        if (fileUrl == "") {
            app.showNotification("The file hasn't been saved yet. Save the file and try again");
        }
        else {
            app.showNotification(fileUrl);
        }
    });
}
```

## <a name="create-a-presentation"></a>Créer une présentation

Votre complément peut créer un nouveau classeur, distinct de l’instance de PowerPoint dans laquelle le complément est en cours d’exécution. L’espace de noms PowerPoint a la `createPresentation` méthode à cet effet. Lorsque cette méthode est appelée, la nouvelle présentation est immédiatement ouverte et affichée dans une nouvelle instance de PowerPoint. Votre complément reste ouvert et en cours d’exécution avec la présentation précédente.

```js
PowerPoint.createPresentation();
```

La `createPresentation` méthode peut également créer une copie d’une présentation existante. La méthode accepte comme un paramètre facultatif une représentation de chaîne codée en base 64 d’un fichier .pptx. La présentation résultante sera une copie de ce fichier, en supposant que l’argument de chaîne est un fichier .pptx valide. La catégorie[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) peut être utilisée pour convertir un fichier dans la chaîne codée en base 64 requise, comme indiqué dans l’exemple suivant.

```js
var myFile = document.getElementById("file");
var reader = new FileReader();

reader.onload = function (event) {
    // strip off the metadata before the base64-encoded string
    var startIndex = reader.result.toString().indexOf("base64,");
    var copyBase64 = reader.result.toString().substr(startIndex + 7);

    PowerPoint.createPresentation(copyBase64);
};

// read in the file as a data URL so we can parse the base64-encoded string
reader.readAsDataURL(myFile.files[0]);
```

## <a name="see-also"></a>Voir aussi

- [Développement de compléments Office](../develop/develop-overview.md)
- [Découvrez le programme pour les développeurs Microsoft 365](https://developer.microsoft.com/microsoft-365/dev-program)
- [Exemples de code PowerPoint](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Utiliser des thèmes de document dans vos compléments PowerPoint](use-document-themes-in-your-powerpoint-add-ins.md)
