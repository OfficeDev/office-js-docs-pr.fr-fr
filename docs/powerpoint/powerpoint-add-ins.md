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
# <a name="powerpoint-add-ins"></a><span data-ttu-id="acaa3-102">Compléments PowerPoint</span><span class="sxs-lookup"><span data-stu-id="acaa3-102">PowerPoint add-ins</span></span>

<span data-ttu-id="acaa3-p101">Vous pouvez utiliser des compléments PowerPoint afin de créer des solutions attrayantes pour les présentations de vos utilisateurs sur toutes les plateformes, y compris Windows, iOS, Office Online et Mac. Vous pouvez créer deux types de compléments PowerPoint :</span><span class="sxs-lookup"><span data-stu-id="acaa3-p101">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iOS, Office Online, and Mac. You can create one of two types of add-ins:</span></span>

- <span data-ttu-id="acaa3-p102">Utilisez des **compléments de contenu** pour ajouter du contenu HTML5 dynamique à vos présentations. Par exemple, consultez le complément [Diagrammes LucidChart pour PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false), qui vous permet d’injecter un diagramme interactif de LucidChart dans votre support de présentation.</span><span class="sxs-lookup"><span data-stu-id="acaa3-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://store.office.com/app.aspx?assetid=WA104380117&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Productivity&homapppos=3&homchv=2&appredirect=false) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="acaa3-p103">Utilisez des **compléments de volet Office** pour faire apparaître des informations de référence ou insérer des données dans la présentation via un service. Par exemple, consultez le complément [Images Shutterstock](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false), qui vous permet d’ajouter des photos professionnelles à votre présentation.</span><span class="sxs-lookup"><span data-stu-id="acaa3-p103">Use **task pane add-ins** to bring in reference information or insert data into the slide via a service. For example, see the [Shutterstock Images](https://store.office.com/app.aspx?assetid=WA104380169&ui=en-US&rs=en-US&ad=US&clickedfilter=OfficeProductFilter%3APowerPoint&productgroup=PowerPoint&homprd=PowerPoint&sourcecorrid=950950b7-aa6c-4766-95fa-e75d37266c21&homappcat=Editor%2527s%2BPicks&homapppos=0&homchv=1&appredirect=false) add-in, which you can use to add professional photos to your presentation.</span></span> 

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="acaa3-109">Scénarios de complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="acaa3-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="acaa3-110">Les exemples de code dans cet article illustrent certaines tâches de base pour le développement de compléments pour PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="acaa3-110">The code examples in the article show you some basic tasks for developing content add-ins for PowerPoint.</span></span> <span data-ttu-id="acaa3-111">Veuillez noter ce qui suit :</span><span class="sxs-lookup"><span data-stu-id="acaa3-111">Please note the following in 2nd_ProjServ_12 Beta 2:</span></span>

- <span data-ttu-id="acaa3-112">Pour afficher des informations, ces exemples utilisent la fonction `app.showNotification`, qui est incluse dans les modèles de projet de compléments Office de Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="acaa3-112">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="acaa3-113">Si vous n’utilisez pas Visual Studio pour développer votre complément, vous devez remplacerez la fonction `showNotification` par votre propre code.</span><span class="sxs-lookup"><span data-stu-id="acaa3-113">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span> 

- <span data-ttu-id="acaa3-114">Plusieurs de ces exemples utilisent également un objet `Globals` qui est déclaré au-delà de l’étendue de ces fonctions en tant que :   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="acaa3-114">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="acaa3-115">Pour utiliser ces exemples, votre projet de complément doit [référencer la bibliothèque Office.js version 1.1 ou ultérieure](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="acaa3-115">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="acaa3-116">Détecter l’affichage actif de la présentation et gérer l’événement ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="acaa3-116">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="acaa3-117">Si vous créez un complément de contenu, vous devrez obtenir la vue active de la présentation et gérer l’événement `ActiveViewChanged` dans le cadre de votre gestionnaire `Office.Initialize`.</span><span class="sxs-lookup"><span data-stu-id="acaa3-117">If you are building a content add-in, you will need to get the presentation's active view and handle the ActiveViewChanged event, as part of your Office.Initialize handler.</span></span> 

> [!NOTE]
> <span data-ttu-id="acaa3-118">Dans PowerPoint Online, l’événement [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) ne se déclenchera jamais du fait que le mode diaporama est traité comme une nouvelle session.</span><span class="sxs-lookup"><span data-stu-id="acaa3-118">In PowerPoint Online, the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event will never fire as Slide Show mode is treated as a new session. In this case, the add-in must fetch the active view on load, as noted below.</span></span> <span data-ttu-id="acaa3-119">Dans ce cas, le complément doit extraire l’affichage actif au chargement, comme illustré dans l’exemple de code suivant.</span><span class="sxs-lookup"><span data-stu-id="acaa3-119">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="acaa3-120">Dans l’exemple de code suivant :</span><span class="sxs-lookup"><span data-stu-id="acaa3-120">In the following code sample:</span></span>

- <span data-ttu-id="acaa3-121">La fonction `getActiveFileView` appelle la méthode [Document.getActiveViewAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getactiveviewasync-options--callback-) afin de renvoyer si la vue actuelle de la présentation est une vue de « modification » (toutes les vues dans lesquelles vous modifiez des diapositives, telles que les vues **Normal** ou **Vue Structure**) ou « lecture » ( **Diaporama** ou **Mode Lecture**).</span><span class="sxs-lookup"><span data-stu-id="acaa3-121">The getFileView function calls the Document.getActiveViewAsync method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as Normal or Outline View) or "read" (Slide Show or Reading View) view.</span></span>

- <span data-ttu-id="acaa3-122">La fonction `registerActiveViewChanged` appelle la méthode [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) afin d’inscrire un gestionnaire pour l’événement [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="acaa3-122">The  `registerActiveViewChanged` function calls the [addHandlerAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) event.</span></span> 


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="acaa3-123">Accéder à une diapositive spécifique dans la présentation</span><span class="sxs-lookup"><span data-stu-id="acaa3-123">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="acaa3-124">Dans l’exemple de code suivant, la fonction `getSelectedRange` appelle la méthode [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) pour obtenir l’objet JSON renvoyé par `asyncResult.value`, qui contient un tableau nommé **slides**.</span><span class="sxs-lookup"><span data-stu-id="acaa3-124">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named **slides**.</span></span> <span data-ttu-id="acaa3-125">Le tableau **slides** contient les ID, titres et des index de la plage sélectionnée de diapositives (ou de la diapositive en cours, si plusieurs diapositives ne sont pas sélectionnés).</span><span class="sxs-lookup"><span data-stu-id="acaa3-125">The **slides** array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="acaa3-126">Il enregistre également l’ID de la première diapositive de la plage sélectionnée dans une variable globale.</span><span class="sxs-lookup"><span data-stu-id="acaa3-126">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="acaa3-127">Dans l’exemple de code suivant, la fonction `goToFirstSlide` appelle la méthode [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) pour accéder à la première diapositive qui a été identifiée par la fonction `getSelectedRange` indiquée précédemment.</span><span class="sxs-lookup"><span data-stu-id="acaa3-127">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="acaa3-128">Naviguer entre les diapositives de la présentation</span><span class="sxs-lookup"><span data-stu-id="acaa3-128">Navigate between slides in the presentation</span></span>

<span data-ttu-id="acaa3-129">Dans l’exemple de code suivant, la fonction `goToSlideByIndex` appelle la méthode **Document.goToByIdAsync** pour passer à la diapositive suivante dans la présentation.</span><span class="sxs-lookup"><span data-stu-id="acaa3-129">The  `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="acaa3-130">Obtenir l’URL de la présentation</span><span class="sxs-lookup"><span data-stu-id="acaa3-130">Get the URL of the presentation</span></span>

<span data-ttu-id="acaa3-131">Dans l’exemple de code suivant, la fonction `getFileUrl` appelle la méthode [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) pour obtenir l’URL du fichier de présentation.</span><span class="sxs-lookup"><span data-stu-id="acaa3-131">The  `getFileUrl` function calls the [Document.getFileProperties](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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



## <a name="see-also"></a><span data-ttu-id="acaa3-132">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="acaa3-132">See also</span></span>
- [<span data-ttu-id="acaa3-133">Exemples de code PowerPoint</span><span class="sxs-lookup"><span data-stu-id="acaa3-133">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/en-us/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="acaa3-134">Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et de volet Office</span><span class="sxs-lookup"><span data-stu-id="acaa3-134">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="acaa3-135">Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="acaa3-135">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="acaa3-136">Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="acaa3-136">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="acaa3-137">Utiliser des thèmes de document dans vos compléments PowerPoint</span><span class="sxs-lookup"><span data-stu-id="acaa3-137">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
    
