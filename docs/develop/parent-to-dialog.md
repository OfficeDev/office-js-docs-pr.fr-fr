---
title: Transmission de données et de messages à une boîte de dialogue à partir de sa page hôte
description: Découvrez comment transmettre des données à une boîte de dialogue à partir de la page hôte à l’aide des API messageChild et DialogParentMessageReceived.
ms.date: 04/16/2020
localization_priority: Normal
ms.openlocfilehash: cd332a58aa79a81aab7cf5a3d247950ce8bc655e
ms.sourcegitcommit: 803587b324fc8038721709d7db5664025cf03c6b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/17/2020
ms.locfileid: "43547056"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a><span data-ttu-id="cc0c3-103">Transmission de données et de messages à une boîte de dialogue à partir de sa page hôte (aperçu)</span><span class="sxs-lookup"><span data-stu-id="cc0c3-103">Passing data and messages to a dialog box from its host page (preview)</span></span>

<span data-ttu-id="cc0c3-104">Votre complément peut envoyer des messages à partir de la [page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) vers une boîte de dialogue à l’aide de la méthode [messageChild](/javascript/api/office/office.dialog#messagechild-message-) de l’objet [Dialog](/javascript/api/office/office.dialog) .</span><span class="sxs-lookup"><span data-stu-id="cc0c3-104">Your add-in can send messages from the [host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) to a dialog box using the [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method of the [Dialog](/javascript/api/office/office.dialog) object.</span></span>

> [!Important]
>
> - <span data-ttu-id="cc0c3-105">Les API décrites dans cet article sont en aperçu.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-105">The APIs described in this article are in preview.</span></span> <span data-ttu-id="cc0c3-106">Elles sont disponibles pour les développeurs dans le cas d’expérimentation ; mais ne doit pas être utilisé dans un complément de production.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-106">They are available to developers for experimentation; but should not be used in a production add-in.</span></span> <span data-ttu-id="cc0c3-107">Tant que cette API n’est pas publiée, utilisez les techniques décrites dans [transmettre les informations à la boîte de dialogue](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) des compléments de production.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-107">Until this API is released, use the techniques described in [Pass information to the dialog box](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) for production add-ins.</span></span>
> - <span data-ttu-id="cc0c3-108">Les API décrites dans cet article nécessitent Office 365 (la version avec abonnement d’Office).</span><span class="sxs-lookup"><span data-stu-id="cc0c3-108">The APIs described in this article require Office 365 (the subscription version of Office).</span></span> <span data-ttu-id="cc0c3-109">Vous devez utiliser la version et le build mensuels les plus récents du canal du programme Insider.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-109">You should use the latest monthly version and build from the Insiders channel.</span></span> <span data-ttu-id="cc0c3-110">Vous devez participer au programme Office Insider pour obtenir cette version.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-110">You need to be an Office Insider to get this version.</span></span> <span data-ttu-id="cc0c3-111">Pour plus d’informations, reportez-vous à [Participez au programme Office Insider](https://insider.office.com).</span><span class="sxs-lookup"><span data-stu-id="cc0c3-111">For more information, see [Be an Office Insider](https://insider.office.com).</span></span> <span data-ttu-id="cc0c3-112">Veuillez noter que lorsqu’une build est basée sur le canal semi-annuel de production, la prise en charge des fonctionnalités d’aperçu est désactivée pour cette version.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-112">Please note that when a build graduates to the production semi-annual channel, support for preview features is turned off for that build.</span></span>
> - <span data-ttu-id="cc0c3-113">Dans l’étape initiale de l’aperçu, les API sont prises en charge dans Excel, PowerPoint et Word ; mais pas dans Outlook.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-113">In the initial stage of the preview, the APIs are supported in Excel, PowerPoint, and Word; but not in Outlook.</span></span>
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a><span data-ttu-id="cc0c3-114">Utiliser `messageChild()` à partir de la page hôte</span><span class="sxs-lookup"><span data-stu-id="cc0c3-114">Use `messageChild()` from the host page</span></span>

<span data-ttu-id="cc0c3-115">Lorsque vous appelez l’API de boîte de dialogue Office pour ouvrir une boîte de dialogue, un objet [Dialog](/javascript/api/office/office.dialog) est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-115">When you call the Office dialog API to open a dialog box, a [Dialog](/javascript/api/office/office.dialog) object is returned.</span></span> <span data-ttu-id="cc0c3-116">Elle doit être assignée à une variable, qui a généralement une portée plus élevée que la méthode [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) , car l’objet sera référencé par d’autres méthodes.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-116">It should be assigned to a variable, which typically has greater scope than the [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) method because the object will be referenced by other methods.</span></span> <span data-ttu-id="cc0c3-117">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="cc0c3-117">The following is an example:</span></span>

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

<span data-ttu-id="cc0c3-118">Cet `Dialog` objet est doté d’une méthode [messageChild](/javascript/api/office/office.dialog#messagechild-message-) qui envoie une chaîne ou des données JSON à la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-118">This `Dialog` object has a [messageChild](/javascript/api/office/office.dialog#messagechild-message-) method that sends any string, or stringified data, to the dialog box.</span></span> <span data-ttu-id="cc0c3-119">Cela déclenche un `DialogParentMessageReceived` événement dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-119">This raises a `DialogParentMessageReceived` event in the dialog box.</span></span> <span data-ttu-id="cc0c3-120">Votre code doit gérer cet événement, comme indiqué dans la section suivante.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-120">Your code should handle this event, as shown in the next section.</span></span>

<span data-ttu-id="cc0c3-121">Imaginez un scénario dans lequel l’interface utilisateur de la boîte de dialogue doit correspondre à la feuille de calcul active et la position de cette feuille de calcul par rapport aux autres feuilles de calcul.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-121">Consider a scenario in which the UI of the dialog should correlate with the currently active worksheet and that worksheet's position relative to the other worksheets.</span></span> <span data-ttu-id="cc0c3-122">Dans l’exemple suivant, `sheetPropertiesChanged` envoie les propriétés de feuille de calcul Excel dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-122">In the following example, `sheetPropertiesChanged` sends Excel worksheet properties to the dialog box.</span></span> <span data-ttu-id="cc0c3-123">Dans ce cas, la feuille de calcul active est nommée « ma feuille » et est la seconde feuille du classeur.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-123">In this case the current worksheet is named "My Sheet" and it is the 2nd sheet in the workbook.</span></span> <span data-ttu-id="cc0c3-124">Les données sont encapsulées dans un objet qui est JSON afin de pouvoir être transmis à `messageChild`.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-124">The data is encapsulated in an object which is stringified so that it can be passed to `messageChild`.</span></span>

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a><span data-ttu-id="cc0c3-125">Gérer DialogParentMessageReceived dans la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="cc0c3-125">Handle DialogParentMessageReceived in the dialog box</span></span>

<span data-ttu-id="cc0c3-126">Dans le JavaScript de la boîte de dialogue, inscrivez un gestionnaire `DialogParentMessageReceived` pour l’événement à l’aide de la méthode [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) .</span><span class="sxs-lookup"><span data-stu-id="cc0c3-126">In the dialog box's JavaScript, register a handler for the `DialogParentMessageReceived` event with the [UI.addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) method.</span></span> <span data-ttu-id="cc0c3-127">Cette opération s’effectue généralement dans les [méthodes Office. onReady ou Office. Initialize](initialize-add-in.md).</span><span class="sxs-lookup"><span data-stu-id="cc0c3-127">This is typically done in the [Office.onReady or Office.initialize methods](initialize-add-in.md).</span></span> <span data-ttu-id="cc0c3-128">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="cc0c3-128">The following is an example:</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

<span data-ttu-id="cc0c3-129">Ensuite, définissez le `onMessageFromParent` gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-129">Then, define the `onMessageFromParent` handler.</span></span> <span data-ttu-id="cc0c3-130">Le code suivant poursuit l’exemple de la section précédente.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-130">The following code continues the example from the preceding section.</span></span> <span data-ttu-id="cc0c3-131">Notez qu’Office transmet un argument au gestionnaire et que la `message` propriété de l’objet argument contient la chaîne de la page hôte.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-131">Note that Office passes an argument to the handler and that the `message` property of argument object contains the string from the host page.</span></span> <span data-ttu-id="cc0c3-132">Dans cet exemple, le message est reconverti en objet et jQuery est utilisé pour définir le titre supérieur de la boîte de dialogue de sorte qu’il corresponde au nouveau nom de la feuille de calcul.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-132">In this example, the message is reconverted to an object and jQuery is used to set the top heading of the dialog to match the new worksheet name.</span></span>

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

<span data-ttu-id="cc0c3-133">Il est recommandé de vérifier que votre gestionnaire est correctement enregistré.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-133">It is a best practice to verify that your handler is properly registered.</span></span> <span data-ttu-id="cc0c3-134">Pour ce faire, vous pouvez transmettre un rappel à `addHandlerAsync` la méthode qui s’exécute lorsque la tentative d’enregistrement du gestionnaire est terminée.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-134">You can do this by passing a callback to the `addHandlerAsync` method that runs when the attempt to register the handler completes.</span></span> <span data-ttu-id="cc0c3-135">Utilisez le gestionnaire pour consigner ou afficher une erreur si le gestionnaire n’a pas été enregistré correctement.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-135">Use the handler to log or show an error if the handler was not successfully registered.</span></span> <span data-ttu-id="cc0c3-136">Voici un exemple.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-136">The following is an example.</span></span> <span data-ttu-id="cc0c3-137">Notez qu' `reportError` il s’agit d’une fonction, non définie ici, qui enregistre ou affiche l’erreur.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-137">Note that `reportError` is a function, not defined here, that logs or displays the error.</span></span>

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a><span data-ttu-id="cc0c3-138">Messagerie conditionnelle</span><span class="sxs-lookup"><span data-stu-id="cc0c3-138">Conditional messaging</span></span>

<span data-ttu-id="cc0c3-139">Étant donné que vous pouvez `messageChild` effectuer plusieurs appels à partir de la page hôte, mais que vous n’avez qu’un seul `DialogParentMessageReceived` gestionnaire dans la boîte de dialogue de l’événement, le gestionnaire doit utiliser une logique conditionnelle pour distinguer les différents messages.</span><span class="sxs-lookup"><span data-stu-id="cc0c3-139">Because you can make multiple `messageChild` calls from the host page, but you have only one handler in the dialog box for the `DialogParentMessageReceived` event, the handler must use conditional logic to distinguish different messages.</span></span> <span data-ttu-id="cc0c3-140">Vous pouvez effectuer cette opération d’une manière parfaitement parallèle à la façon dont vous structurez la messagerie conditionnelle lorsque la boîte de dialogue envoie un message à la page hôte, comme décrit dans la section [messagerie conditionnelle](dialog-api-in-office-add-ins.md#conditional-messaging).</span><span class="sxs-lookup"><span data-stu-id="cc0c3-140">You can do this in a way that is precisely parallel to how you would structure conditional messaging when the dialog box is sending a message to the host page as described in [Conditional messaging](dialog-api-in-office-add-ins.md#conditional-messaging).</span></span>
