---
title: Compléments PowerPoint
description: ''
ms.date: 11/07/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 798e7c8357f2f2f021a76c0b88b0617764428c8c
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851340"
---
# <a name="powerpoint-add-ins"></a><span data-ttu-id="c662d-102">Compléments PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c662d-102">PowerPoint add-ins</span></span>

<span data-ttu-id="c662d-103">Vous pouvez utiliser des compléments PowerPoint afin de créer des solutions attrayantes pour les présentations de vos utilisateurs sur différentes plateformes, notamment Windows, iPad et Mac, ainsi que dans un navigateur.</span><span class="sxs-lookup"><span data-stu-id="c662d-103">You can use PowerPoint add-ins to build engaging solutions for your users' presentations across platforms including Windows, iPad, Mac, and in a browser.</span></span> <span data-ttu-id="c662d-104">Vous pouvez créer deux types de commandes de complément PowerPoint:</span><span class="sxs-lookup"><span data-stu-id="c662d-104">You can create two types of PowerPoint add-ins:</span></span>

- <span data-ttu-id="c662d-p102">Utilisez des **compléments de contenu** pour ajouter du contenu HTML5 dynamique à vos présentations. Par exemple, consultez le complément [Diagrammes LucidChart pour PowerPoint](https://appsource.microsoft.com/product/office/WA104380117), qui vous permet d’injecter un diagramme interactif de LucidChart dans votre support de présentation.</span><span class="sxs-lookup"><span data-stu-id="c662d-p102">Use **content add-ins** to add dynamic HTML5 content to your presentations. For example, see the [LucidChart Diagrams for PowerPoint](https://appsource.microsoft.com/product/office/WA104380117) add-in, which you can use to inject an interactive diagram from LucidChart into your deck.</span></span>

- <span data-ttu-id="c662d-107">Utilisez des **compléments de volet Office** pour faire apparaître des informations de référence ou insérer des données dans la diapositive via un service.</span><span class="sxs-lookup"><span data-stu-id="c662d-107">Use **task pane add-ins** to bring in reference information or insert data into the presentation via a service.</span></span> <span data-ttu-id="c662d-108">Par exemple, consultez le complément [Pixton Comic Characters](https://appsource.microsoft.com/product/office/WA104380907), qui vous permet d’ajouter des photos professionnelles à votre présentation.</span><span class="sxs-lookup"><span data-stu-id="c662d-108">For example, see the [Pixton Comic Characters](https://appsource.microsoft.com/product/office/WA104380907) add-in, which you can use to add professional photos to your presentation.</span></span>

## <a name="powerpoint-add-in-scenarios"></a><span data-ttu-id="c662d-109">Scénarios de complément PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c662d-109">PowerPoint add-in scenarios</span></span>

<span data-ttu-id="c662d-110">Les exemples de code figurant dans l’article vous présentent certaines tâches de base en matière de développement de compléments de contenu pour PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="c662d-110">The code examples in this article demonstrate some basic tasks for developing add-ins for PowerPoint.</span></span> <span data-ttu-id="c662d-111">Notez également ce qui suit:</span><span class="sxs-lookup"><span data-stu-id="c662d-111">Please note the following:</span></span>

- <span data-ttu-id="c662d-112">Pour afficher des informations, ces exemples dépendent de la fonction`app.showNotification`, qui est incluse dans les modèles de projet de compléments Office Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="c662d-112">To display information, these examples use the `app.showNotification` function, which is included in the Visual Studio Office Add-ins project templates.</span></span> <span data-ttu-id="c662d-113">Si vous n’utilisez pas Visual Studio pour développer votre complément, vous devrez remplacer la fonction`showNotification`par votre propre code.</span><span class="sxs-lookup"><span data-stu-id="c662d-113">If you aren't using Visual Studio to develop your add-in, you'll need replace the `showNotification` function with your own code.</span></span>

- <span data-ttu-id="c662d-114">Plusieurs de ces exemples dépendent également de l’objet`Globals` qui est déclaré en dehors de la portée de ces fonctions: `var Globals = {activeViewHandler:0, firstSlideId:0};`</span><span class="sxs-lookup"><span data-stu-id="c662d-114">Several of these examples also use a `Globals` object that is declared beyond the scope of these functions as:   `var Globals = {activeViewHandler:0, firstSlideId:0};`</span></span>

- <span data-ttu-id="c662d-115">Pour utiliser ces exemples, votre projet complément doit [référencer Office.js version 1.1 bibliothèque ou version ultérieure](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="c662d-115">To use these examples, your add-in project must [reference Office.js v1.1 library or later](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="detect-the-presentations-active-view-and-handle-the-activeviewchanged-event"></a><span data-ttu-id="c662d-116">Détecter l’affichage actif de la présentation et gérer l’événement ActiveViewChanged</span><span class="sxs-lookup"><span data-stu-id="c662d-116">Detect the presentation's active view and handle the ActiveViewChanged event</span></span>

<span data-ttu-id="c662d-117">Si vous créez un complément de contenu, vous devrez obtenir la vue active de la présentation et gérer`ActiveViewChanged`l’événement ActiveViewChanged dans le cadre de votre`Office.Initialize`gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="c662d-117">If you are building a content add-in, you will need to get the presentation's active view and handle the `ActiveViewChanged` event, as part of your `Office.Initialize` handler.</span></span>

> [!NOTE]
> <span data-ttu-id="c662d-118">Dans PowerPoint sur le web, l’événement [Document.ActiveViewChanged](/javascript/api/office/office.document) ne se déclenche jamais, car le mode Diaporama est considéré comme une nouvelle session.</span><span class="sxs-lookup"><span data-stu-id="c662d-118">In PowerPoint on the web, the [Document.ActiveViewChanged](/javascript/api/office/office.document) event will never fire as Slide Show mode is treated as a new session.</span></span> <span data-ttu-id="c662d-119">Dans ce cas, le complément doit extraire la vue active lors du chargement, comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="c662d-119">In this case, the add-in must fetch the active view on load, as shown in the following code sample.</span></span>

<span data-ttu-id="c662d-120">Collez le code suivant:</span><span class="sxs-lookup"><span data-stu-id="c662d-120">In the following code sample:</span></span>

- <span data-ttu-id="c662d-121">La fonction`getActiveFileView` appelle la méthode [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) afin de renvoyer si la vue actuelle de la présentation est une vue de « modification » (toutes les vues dans lesquelles vous modifiez des diapositives, telles que les vues **Normal** ou **Mode Plan**) ou « lecture » ( **Diaporama**ou**Mode Lecture**).</span><span class="sxs-lookup"><span data-stu-id="c662d-121">The  `getActiveFileView` function calls the [Document.getActiveViewAsync](/javascript/api/office/office.document#getactiveviewasync-options--callback-) method to return whether the presentation's current view is "edit" (any of the views in which you can edit slides, such as **Normal** or **Outline View**) or "read" (**Slide Show** or **Reading View**).</span></span>

- <span data-ttu-id="c662d-122">La fonction`registerActiveViewChanged`appelle la méthode [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) afin d’inscrire un gestionnaire pour l’événement [Document.ActiveViewChanged](/javascript/api/office/office.document).</span><span class="sxs-lookup"><span data-stu-id="c662d-122">The  `registerActiveViewChanged` function calls the [addHandlerAsync](/javascript/api/office/office.document#addhandlerasync-eventtype--handler--options--callback-) method to register a handler for the [Document.ActiveViewChanged](/javascript/api/office/office.document) event.</span></span>


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

## <a name="navigate-to-a-particular-slide-in-the-presentation"></a><span data-ttu-id="c662d-123">Accéder à une diapositive spécifique dans la présentation</span><span class="sxs-lookup"><span data-stu-id="c662d-123">Navigate to a particular slide in the presentation</span></span>

<span data-ttu-id="c662d-124">Dans l’exemple de code suivant, la fonction`getSelectedRange`appelle la méthode[Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-)pour obtenir l’objet JSON renvoyées par`asyncResult.value`, qui contient un tableau nommé **diapositives**.</span><span class="sxs-lookup"><span data-stu-id="c662d-124">In the following code sample, the `getSelectedRange` function calls the [Document.getSelectedDataAsync](/javascript/api/office/office.document#getselecteddataasync-coerciontype--options--callback-) method to get the JSON object returned by `asyncResult.value`, which contains an array named **slides**.</span></span> <span data-ttu-id="c662d-125">La matrice**diapositives**contient les IDs, les titres et les indexes de plage sélectionnées de diapositives (ou de la diapositive active si plusieurs diapositives ne sont pas sélectionnées).</span><span class="sxs-lookup"><span data-stu-id="c662d-125">The **slides** array contains the ids, titles, and indexes of selected range of slides (or of the current slide, if multiple slides are not selected).</span></span> <span data-ttu-id="c662d-126">Elle enregistre également l’id de la première diapositive dans la plage sélectionnée à une variable globale.</span><span class="sxs-lookup"><span data-stu-id="c662d-126">It also saves the id of the first slide in the selected range to a global variable.</span></span>

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

<span data-ttu-id="c662d-127">Dans l’exemple de code suivant la fonction`goToFirstSlide`appelle la méthode[Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-)pour accéder à la première diapositive qui a été identifiée par la fonction`getSelectedRange`illustrée précédemment.</span><span class="sxs-lookup"><span data-stu-id="c662d-127">In the following code sample, the `goToFirstSlide` function calls the [Document.goToByIdAsync](/javascript/api/office/office.document#gotobyidasync-id--gototype--options--callback-) method to navigate to the first slide that was identified by the `getSelectedRange` function shown previously.</span></span>

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

## <a name="navigate-between-slides-in-the-presentation"></a><span data-ttu-id="c662d-128">Naviguer entre les diapositives de la présentation</span><span class="sxs-lookup"><span data-stu-id="c662d-128">Navigate between slides in the presentation</span></span>

<span data-ttu-id="c662d-129">La fonction`goToSlideByIndex` appelle la méthode **Document.goToByIdAsync** pour passer à la diapositive suivante dans la présentation.</span><span class="sxs-lookup"><span data-stu-id="c662d-129">In the following code sample, the `goToSlideByIndex` function calls the **Document.goToByIdAsync** method to navigate to the next slide in the presentation.</span></span>

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

## <a name="get-the-url-of-the-presentation"></a><span data-ttu-id="c662d-130">Obtenir l’URL de la présentation</span><span class="sxs-lookup"><span data-stu-id="c662d-130">Get the URL of the presentation</span></span>

<span data-ttu-id="c662d-131">La fonction`getFileUrl` appelle la méthode [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) pour obtenir l’URL du fichier de présentation.</span><span class="sxs-lookup"><span data-stu-id="c662d-131">In the following code sample, the  `getFileUrl` function calls the [Document.getFileProperties](/javascript/api/office/office.document#getfilepropertiesasync-options--callback-) method to get the URL of the presentation file.</span></span>

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

## <a name="create-a-presentation"></a><span data-ttu-id="c662d-132">Créer une présentation</span><span class="sxs-lookup"><span data-stu-id="c662d-132">Create a presentation</span></span>

<span data-ttu-id="c662d-133">Votre complément peut créer un nouveau classeur, distinct de l’instance de PowerPoint dans laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="c662d-133">Your add-in can create a new presentation, separate from the PowerPoint instance in which the add-in is currently running.</span></span> <span data-ttu-id="c662d-134">L’espace de noms PowerPoint a la `createPresentation` méthode à cet effet.</span><span class="sxs-lookup"><span data-stu-id="c662d-134">The PowerPoint namespace has the `createPresentation` method for this purpose.</span></span> <span data-ttu-id="c662d-135">Lorsque cette méthode est appelée, la nouvelle présentation est immédiatement ouverte et affichée dans une nouvelle instance de PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="c662d-135">When this method is called, the new presentation is immediately opened and displayed in a new instance of PowerPoint.</span></span> <span data-ttu-id="c662d-136">Votre complément reste ouvert et en cours d’exécution avec la présentation précédente.</span><span class="sxs-lookup"><span data-stu-id="c662d-136">Your add-in remains open and running with the previous presentation.</span></span>

```js
PowerPoint.createPresentation();
```

<span data-ttu-id="c662d-137">La `createPresentation` méthode peut également créer une copie d’une présentation existante.</span><span class="sxs-lookup"><span data-stu-id="c662d-137">The `createPresentation` method can also create a copy of an existing presentation.</span></span> <span data-ttu-id="c662d-138">La méthode accepte comme un paramètre facultatif une représentation de chaîne codée en base 64 d’un fichier .pptx.</span><span class="sxs-lookup"><span data-stu-id="c662d-138">The method accepts a base64-encoded string representation of an .pptx file as an optional parameter.</span></span> <span data-ttu-id="c662d-139">La présentation résultante sera une copie de ce fichier, en supposant que l’argument de chaîne est un fichier .pptx valide.</span><span class="sxs-lookup"><span data-stu-id="c662d-139">The resulting presentation will be a copy of that file, assuming the string argument is a valid .pptx file.</span></span> <span data-ttu-id="c662d-140">La catégorie[FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) peut être utilisée pour convertir un fichier dans la chaîne codée en base 64 requise, comme indiqué dans l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="c662d-140">The [FileReader](https://developer.mozilla.org/docs/Web/API/FileReader) class can be used to convert a file into the required base64-encoded string, as demonstrated in the following example.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="c662d-141">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="c662d-141">See also</span></span>

- [<span data-ttu-id="c662d-142">Création de compléments Office</span><span class="sxs-lookup"><span data-stu-id="c662d-142">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="c662d-143">Exemples de code PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c662d-143">PowerPoint Code Samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,PowerPoint)
- [<span data-ttu-id="c662d-144">Enregistrement de l’état et des paramètres d’un complément par document pour les compléments de contenu et du volet Office</span><span class="sxs-lookup"><span data-stu-id="c662d-144">How to save add-in state and settings per document for content and task pane add-ins</span></span>](../develop/persisting-add-in-state-and-settings.md#how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins)
- [<span data-ttu-id="c662d-145">Lecture et écriture de données dans la sélection active d’un document ou d’une feuille de calcul</span><span class="sxs-lookup"><span data-stu-id="c662d-145">Read and write data to the active selection in a document or spreadsheet</span></span>](../develop/read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md)
- [<span data-ttu-id="c662d-146">Obtention de l’intégralité d’un document pour un complément pour PowerPoint ou Word</span><span class="sxs-lookup"><span data-stu-id="c662d-146">Get the whole document from an add-in for PowerPoint or Word</span></span>](../powerpoint/get-the-whole-document-from-an-add-in-for-powerpoint.md)
- [<span data-ttu-id="c662d-147">Utiliser des thèmes de document dans vos compléments PowerPoint</span><span class="sxs-lookup"><span data-stu-id="c662d-147">Use document themes in your PowerPoint add-ins</span></span>](use-document-themes-in-your-powerpoint-add-ins.md)
