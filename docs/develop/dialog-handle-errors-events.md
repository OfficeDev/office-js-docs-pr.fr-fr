---
title: Gestion des erreurs et des événements dans la boîte de dialogue Office
description: Décrit comment éviter et gérer les erreurs lors de l’ouverture et de l’utilisation de la boîte Office dialogue
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: be1fb8bcd30b47ac6399657d928d3cad7f857f39
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349895"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="1dc37-103">Gestion des erreurs et des événements dans la boîte de dialogue Office</span><span class="sxs-lookup"><span data-stu-id="1dc37-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="1dc37-104">Cet article explique comment prendre en charge les erreurs lors de l’ouverture de la boîte de dialogue et les erreurs qui se produisent à l’intérieur de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1dc37-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="1dc37-105">Cet article présuppose que vous connaissez les principes de base de l’utilisation de l’API de boîte de dialogue Office, comme décrit dans [l’API](dialog-api-in-office-add-ins.md)de boîte de dialogue Office dans vos Office.</span><span class="sxs-lookup"><span data-stu-id="1dc37-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="1dc37-106">Voir aussi [Meilleures pratiques et règles pour l’API Office boîte de dialogue .](dialog-best-practices.md)</span><span class="sxs-lookup"><span data-stu-id="1dc37-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="1dc37-107">Votre code doit gérer deux catégories d’événements :</span><span class="sxs-lookup"><span data-stu-id="1dc37-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="1dc37-108">les erreurs renvoyées par l’appel de `displayDialogAsync` car la boîte de dialogue ne peut pas être créée ;</span><span class="sxs-lookup"><span data-stu-id="1dc37-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="1dc37-109">Erreurs et autres événements dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1dc37-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="1dc37-110">Erreurs provenant de displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="1dc37-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="1dc37-111">Outre les erreurs générales de plateforme et système, quatre erreurs sont spécifiques à `displayDialogAsync` l’appel.</span><span class="sxs-lookup"><span data-stu-id="1dc37-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="1dc37-112">Numéro de code</span><span class="sxs-lookup"><span data-stu-id="1dc37-112">Code number</span></span>|<span data-ttu-id="1dc37-113">Signification</span><span class="sxs-lookup"><span data-stu-id="1dc37-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="1dc37-114">12004</span><span class="sxs-lookup"><span data-stu-id="1dc37-114">12004</span></span>|<span data-ttu-id="1dc37-p101">Le domaine de l’URL transmis à `displayDialogAsync` n’est pas approuvé. Le domaine doit être le même domaine que celui de la page hôte (y compris le protocole et le numéro de port).</span><span class="sxs-lookup"><span data-stu-id="1dc37-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="1dc37-117">12005</span><span class="sxs-lookup"><span data-stu-id="1dc37-117">12005</span></span>|<span data-ttu-id="1dc37-118">L’URL transmise à `displayDialogAsync` utilise le protocole HTTP.</span><span class="sxs-lookup"><span data-stu-id="1dc37-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="1dc37-119">C’est le protocole HTTPS qui est requis.</span><span class="sxs-lookup"><span data-stu-id="1dc37-119">HTTPS is required.</span></span> <span data-ttu-id="1dc37-120">(Dans certaines versions de Office, le texte du message d’erreur renvoyé avec 12005 est identique à celui renvoyé pour 12004.)</span><span class="sxs-lookup"><span data-stu-id="1dc37-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="1dc37-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="1dc37-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="1dc37-p103">Une boîte de dialogue est déjà ouverte à partir de cette fenêtre hôte. Une fenêtre hôte, par exemple un volet Office, ne peut avoir qu’une seule boîte de dialogue ouverte à la fois.</span><span class="sxs-lookup"><span data-stu-id="1dc37-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="1dc37-124">12009</span><span class="sxs-lookup"><span data-stu-id="1dc37-124">12009</span></span>|<span data-ttu-id="1dc37-125">L’utilisateur a choisi d’ignorer la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1dc37-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="1dc37-126">Cette erreur peut se produire dans Office sur le Web, où les utilisateurs peuvent choisir de ne pas autoriser un add-in à présenter une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="1dc37-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="1dc37-127">Pour plus d’informations, voir [Gestion des bloqueurs de](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web)fenêtres Office sur le Web .</span><span class="sxs-lookup"><span data-stu-id="1dc37-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="1dc37-128">Lorsqu’elle est appelée, elle transmet un `displayDialogAsync` objet [AsyncResult](/javascript/api/office/office.asyncresult) à sa fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="1dc37-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="1dc37-129">Lorsque l’appel réussit, la boîte de dialogue s’ouvre et la propriété de l’objet `value` `AsyncResult` est un objet [Dialog.](/javascript/api/office/office.dialog)</span><span class="sxs-lookup"><span data-stu-id="1dc37-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="1dc37-130">Pour obtenir un exemple de cela, voir Envoyer des informations à partir de la boîte [de dialogue vers la page hôte.](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page)</span><span class="sxs-lookup"><span data-stu-id="1dc37-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="1dc37-131">Lorsque l’appel échoue, la boîte de dialogue n’est pas créée, la propriété de l’objet est définie sur et la propriété de `displayDialogAsync` `status` `AsyncResult` l’objet est `Office.AsyncResultStatus.Failed` `error` remplie.</span><span class="sxs-lookup"><span data-stu-id="1dc37-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="1dc37-132">Vous devez toujours fournir un rappel qui teste et répond en cas `status` d’erreur.</span><span class="sxs-lookup"><span data-stu-id="1dc37-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="1dc37-133">Pour obtenir un exemple qui signale le message d’erreur, quel que soit son numéro de code, consultez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="1dc37-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="1dc37-134">(La `showNotification` fonction, non définie dans cet article, affiche ou enregistre l’erreur.</span><span class="sxs-lookup"><span data-stu-id="1dc37-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="1dc37-135">Pour obtenir un exemple de la façon dont vous pouvez implémenter cette fonction dans votre application, voir Office’API de boîte de dialogue [de l’application.)](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)</span><span class="sxs-lookup"><span data-stu-id="1dc37-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
function (asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        showNotification(asyncResult.error.code = ": " + asyncResult.error.message);
    } else {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
});
```

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="1dc37-136">Erreurs et événements dans la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="1dc37-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="1dc37-137">Trois erreurs et événements dans la boîte de dialogue lèvent un `DialogEventReceived` événement dans la page hôte.</span><span class="sxs-lookup"><span data-stu-id="1dc37-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="1dc37-138">Pour un rappel de ce qu’est une page hôte, voir Ouvrir une boîte de [dialogue à partir d’une page hôte.](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)</span><span class="sxs-lookup"><span data-stu-id="1dc37-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="1dc37-139">Numéro de code</span><span class="sxs-lookup"><span data-stu-id="1dc37-139">Code number</span></span>|<span data-ttu-id="1dc37-140">Signification</span><span class="sxs-lookup"><span data-stu-id="1dc37-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="1dc37-141">12002</span><span class="sxs-lookup"><span data-stu-id="1dc37-141">12002</span></span>|<span data-ttu-id="1dc37-142">Un des éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="1dc37-142">One of the following:</span></span><br> <span data-ttu-id="1dc37-143">- Aucune page n’existe à l’URL qui a été transmise à `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="1dc37-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="1dc37-144">- Page qui a été transmise au chargement, mais la boîte de dialogue a ensuite été redirigée vers une page qu’elle ne peut ni trouver ni charger, ou qui a été redirigée vers une URL dont la syntaxe n’est `displayDialogAsync` pas valide.</span><span class="sxs-lookup"><span data-stu-id="1dc37-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="1dc37-145">12003</span><span class="sxs-lookup"><span data-stu-id="1dc37-145">12003</span></span>|<span data-ttu-id="1dc37-p107">La boîte de dialogue a été redirigée vers une URL avec le protocole HTTP. C’est le protocole HTTPS qui est requis.</span><span class="sxs-lookup"><span data-stu-id="1dc37-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="1dc37-148">12006</span><span class="sxs-lookup"><span data-stu-id="1dc37-148">12006</span></span>|<span data-ttu-id="1dc37-149">La boîte de dialogue a été fermée, généralement parce que l’utilisateur a choisi **le bouton** **Fermer X**.</span><span class="sxs-lookup"><span data-stu-id="1dc37-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="1dc37-p108">Votre code peut attribuer un gestionnaire pour l’événement `DialogEventReceived` dans l’appel de `displayDialogAsync`. Voici un exemple simple.</span><span class="sxs-lookup"><span data-stu-id="1dc37-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example.</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="1dc37-152">Pour obtenir un exemple de handler pour l’événement qui crée des messages d’erreur personnalisés pour chaque `DialogEventReceived` code d’erreur, voir l’exemple suivant.</span><span class="sxs-lookup"><span data-stu-id="1dc37-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example.</span></span>

```js
function processDialogEvent(arg) {
    switch (arg.error) {
        case 12002:
            showNotification("The dialog box has been directed to a page that it cannot find or load, or the URL syntax is invalid.");
            break;
        case 12003:
            showNotification("The dialog box has been directed to a URL with the HTTP protocol. HTTPS is required.");            break;
        case 12006:
            showNotification("Dialog closed.");
            break;
        default:
            showNotification("Unknown error in dialog box.");
            break;
    }
}
```

<span data-ttu-id="1dc37-153">Pour voir un exemple de complément qui gère les erreurs de cette façon, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="1dc37-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>
