---
title: Compléments PowerPoint
description: ''
ms.date: 10/16/2018
ms.openlocfilehash: 390497e74d4dc52b9d400f242850ab72bdb0eabc
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640077"
---
# <a name="powerpoint-add-ins"></a>Compléments PowerPoint

Vous pouvez utiliser des compléments PowerPoint afin de créer des solutions attrayantes pour les présentations de vos utilisateurs sur toutes les plateformes, y compris Windows, iOS, Office Online et Mac. Vous pouvez créer deux types de compléments PowerPoint :

- Utilisez des **compléments de contenu** pour ajouter du contenu HTML5 dynamique à vos présentations. Par exemple, consultez le complément [Diagrammes LucidChart pour PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), qui vous permet d’injecter un diagramme interactif de LucidChart dans votre support de présentation.

- Utilisez des **compléments de volet Office** pour faire apparaître des informations de référence ou insérer des données dans la présentation via un service. Par exemple, consultez le complément [Images Shutterstock](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), qui vous permet d’ajouter des photos professionnelles à votre présentation. 

## <a name="powerpoint-add-in-scenarios"></a>Scénarios de complément PowerPoint

Les exemples de code dans cet article illustrent certaines tâches de base pour le développement de compléments pour PowerPoint. Veuillez noter ce qui suit :

- Pour afficher des informations, ces exemples utilisent la fonction `app.showNotification`, qui est incluse dans les modèles de projet de compléments Office de Visual Studio. Si vous n’utilisez pas Visual Studio pour développer votre complément, vous devez remplacerez la fonction `showNotification` par votre propre code. 

- Plusieurs de ces exemples utilisent également un objet `Globals` qui est déclaré au-delà de l’étendue de ces fonctions en tant que :   `var Globals = {activeViewHandler:0, firstSlideId:0};`

- Pour utiliser ces exemples, votre projet de complément doit [référencer la bibliothèque Office.js version 1.1 ou ultérieure](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a>Détecter l’affichage actif de la présentation et gérer l’événement ActiveViewChanged

Si vous créez un complément de contenu, vous devrez obtenir la vue active de la présentation et gérer l’événement `ActiveViewChanged` dans le cadre de votre gestionnaire `Office.Initialize`. 

> [!NOTE]
> Dans PowerPoint Online, l’événement [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) ne se déclenchera jamais du fait que le mode diaporama est traité comme une nouvelle session. Dans ce cas, le complément doit extraire l’affichage actif au chargement, comme illustré dans l’exemple de code suivant.

Dans l’exemple de code suivant :

- La fonction `getActiveFileView` appelle la méthode [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) afin de renvoyer si la vue actuelle de la présentation est une vue de « modification » (toutes les vues dans lesquelles vous modifiez des diapositives, telles que les vues **Normal** ou **Vue Structure**) ou « lecture » ( **Diaporama** ou **Mode Lecture**).

- La fonction `registerActiveViewChanged` appelle la méthode [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) afin d’inscrire un gestionnaire pour l’événement [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js). 


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

Dans l’exemple de code suivant, la fonction `getSelectedRange` appelle la méthode [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) pour obtenir l’objet JSON renvoyé par `asyncResult.value`, qui contient un tableau nommé **slides**. Le tableau **slides** contient les ID, titres et des index de la plage sélectionnée de diapositives (ou de la diapositive en cours, si plusieurs diapositives ne sont pas sélectionnés). Il enregistre également l’ID de la première diapositive de la plage sélectionnée dans une variable globale.

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

Dans l’exemple de code suivant, la fonction `goToFirstSlide` appelle la méthode [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) pour accéder à la première diapositive qui a été identifiée par la fonction `getSelectedRange` indiquée précédemment.

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

Dans l’exemple de code suivant, la fonction `goToSlideByIndex` appelle la méthode **Document.goToByIdAsync** pour passer à la diapositive suivante dans la présentation.

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

Dans l’exemple de code suivant, la fonction `getFileUrl` appelle la méthode [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) pour obtenir l’URL du fichier de présentation.

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



## <a name="see-also"></a>Voir aussi
- [Exemples de code PowerPoint](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et de volet Office](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul.](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [Utiliser des thèmes de document dans vos compléments PowerPoint](use-document-themes-in-your-powerpoint-add-ins.md)
    
