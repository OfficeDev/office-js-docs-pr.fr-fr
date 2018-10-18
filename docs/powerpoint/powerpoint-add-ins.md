---
title: Compléments PowerPoint
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 21f6ec0b00003a90df6850562e399d33da7b49b9
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2018
ms.locfileid: "23943886"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="64720-102">Compléments PowerPoint</span><span class="sxs-lookup"><span data-stu-id="64720-102">PowerPoint add-ins</span></span>

<span data-ttu-id="64720-p101">Vous pouvez utiliser des compléments PowerPoint afin de créer des solutions attrayantes pour les présentations de vos utilisateurs sur toutes les plateformes, y compris Windows, iOS, Office Online et Mac. Vous pouvez créer l’un des deux types de compléments :</span><span class="sxs-lookup"><span data-stu-id="64720-p101">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span>

- <span data-ttu-id="64720-p102">Utilisez des **compléments de contenu** pour ajouter du contenu HTML5 dynamique à vos présentations. Par exemple, consultez le complément [Diagrammes LucidChart pour PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), qui vous permet d’injecter un diagramme interactif de LucidChart dans votre support de présentation.</span><span class="sxs-lookup"><span data-stu-id="64720-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>
- <span data-ttu-id="64720-p103">Utilisez des **compléments de volet Office** pour faire apparaître des informations de référence ou insérer des données dans la diapositive via un service. Par exemple, consultez le complément [Images Shutterstock](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), qui vous permet d’ajouter des photos professionnelles à votre présentation.</span><span class="sxs-lookup"><span data-stu-id="64720-p103">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="64720-109">Scénarios de complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="64720-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="64720-110">Les exemples de code figurant dans l’article vous présentent certaines tâches de base en matière de développement de compléments de contenu pour PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="64720-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> 

<span data-ttu-id="64720-p104">Pour afficher des informations, ces exemples dépendent de la fonction `app.showNotification`, qui est incluse dans les modèles de projet de compléments Office Visual Studio. Si vous n’utilisez pas Visual Studio pour développer votre complément, vous devrez remplacer la fonction `showNotification` par votre propre code. Plusieurs de ces exemples dépendent également de l’objet `globals` qui est déclaré en dehors de la portée de ces fonctions : `var globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="64720-p104">To display information, these examples depend on the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates. If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code. Several of these examples also depend on this `globals` object that is declared outside of the scope of these functions: `var globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

<span data-ttu-id="64720-114">Pour obtenir ces exemples de code, votre projet doit faire référence à la [bibliothèque Office.js v1.1 ou version ultérieure](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="64720-114">These code examples require your project to [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>


## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="64720-115">Détecter l’affichage actif de la présentation et gérer l’événement ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="64720-115">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="64720-116">Si vous créez un complément de contenu, vous devrez obtenir la vue active de la présentation et gérer l’événement ActiveViewChanged dans le cadre de votre gestionnaire Office.Initialize.</span><span class="sxs-lookup"><span data-stu-id="64720-116">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span>


- <span data-ttu-id="64720-117">La fonction  `getActiveFileView` appelle la méthode [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) afin de renvoyer si la vue actuelle de la présentation est une vue de « modification » (toutes les vues dans lesquelles vous modifiez des diapositives, telles que les vues **Normal** ou **Mode Plan**) ou « lecture » ( **Diaporama** ou **Mode Lecture**).</span><span class="sxs-lookup"><span data-stu-id="64720-117">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" ( **Slide Show** or **Reading View**) view.</span></span>


- <span data-ttu-id="64720-118">La fonction `registerActiveViewChanged` appelle la méthode [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) afin d’inscrire un gestionnaire pour l’événement [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="64720-118">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event.</span></span> 

> [!NOTE]
> <span data-ttu-id="64720-p105">Dans PowerPoint Online, l’événement [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) ne se déclenche jamais, car le mode diaporama est considéré comme une nouvelle session. Dans ce cas, le complément doit extraire la vue active lors du chargement, comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="64720-p105">In PowerPoint Online, the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span>

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
    

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="64720-121">Accéder à une diapositive spécifique dans la présentation</span><span class="sxs-lookup"><span data-stu-id="64720-121">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="64720-p106">La fonction  `getSelectedRange` appelle la méthode [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) pour obtenir un objet JSON renvoyé par `asyncResult.value` et qui contient un tableau intitulé « diapositives » répertoriant les ID, les titres et les index de la série de diapositives sélectionnée (ou uniquement de la diapositive en cours). Elle enregistre également l’ID de la première diapositive de la série sélectionnée dans une variable globale.</span><span class="sxs-lookup"><span data-stu-id="64720-p106">The  `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method to get a JSON object returned by `asyncResult.value`, which contains an array named "slides" that contains the ids, titles, and indexes of selected range of slides (or just the current slide). It also saves the id of the first slide in the selected range to a global variable.</span></span>


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

<span data-ttu-id="64720-124">La fonction `goToFirstSlide` appelle la méthode [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) pour accéder à l’ID de la première diapositive stockée par la fonction `getSelectedRange` ci-dessus.</span><span class="sxs-lookup"><span data-stu-id="64720-124">The  `goToFirstSlide` function calls the [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) method to go to the id of the first slide stored by the `getSelectedRange` function above.</span></span>




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


## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="64720-125">Naviguer entre les diapositives de la présentation</span><span class="sxs-lookup"><span data-stu-id="64720-125">Navigate between slides in the presentation</span></span>

<span data-ttu-id="64720-126">La fonction `goToSlideByIndex` appelle la méthode **Document.goToByIdAsync** pour passer à la diapositive suivante dans la présentation.</span><span class="sxs-lookup"><span data-stu-id="64720-126">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>


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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="64720-127">Obtenir l’URL de la présentation</span><span class="sxs-lookup"><span data-stu-id="64720-127">Get the URL of the presentation</span></span>

<span data-ttu-id="64720-128">La fonction `getFileUrl` appelle la méthode [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) pour obtenir l’URL du fichier de présentation.</span><span class="sxs-lookup"><span data-stu-id="64720-128">The  `getFileUrl` function calls the [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>


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



## <a name="see-also"></a><span data-ttu-id="64720-129">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="64720-129">See also</span></span>
- [<span data-ttu-id="64720-130">Exemples de code PowerPoint</span><span class="sxs-lookup"><span data-stu-id="64720-130">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="64720-131">Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office</span><span class="sxs-lookup"><span data-stu-id="64720-131">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="64720-132">Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="64720-132">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="64720-133">Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="64720-133">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="64720-134">Utiliser des thèmes de document dans vos compléments PowerPoint</span><span class="sxs-lookup"><span data-stu-id="64720-134">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
