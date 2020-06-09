---
title: Gestion des erreurs et des événements dans la boîte de dialogue Office
description: Indique comment intercepter et gérer les erreurs lors de l’ouverture et de l’utilisation de la boîte de dialogue Office
ms.date: 01/29/2020
localization_priority: Normal
ms.openlocfilehash: d83d5c4627f68c3f4b1c196cf543d01bf981abbe
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608173"
---
# <a name="handling-errors-and-events-in-the-office-dialog-box"></a><span data-ttu-id="0297f-103">Gestion des erreurs et des événements dans la boîte de dialogue Office</span><span class="sxs-lookup"><span data-stu-id="0297f-103">Handling errors and events in the Office dialog box</span></span>

<span data-ttu-id="0297f-104">Cet article explique comment intercepter et gérer les erreurs lors de l’ouverture de la boîte de dialogue et des erreurs qui se produisent dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="0297f-104">This article describes how to trap and handle errors when opening the dialog box and errors that happen inside the dialog box.</span></span>

> [!NOTE]
> <span data-ttu-id="0297f-105">Cet article suppose que vous êtes familiarisé avec les notions de base de l’utilisation de l’API de boîte de dialogue Office, comme décrit dans la rubrique [use the Office Dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="0297f-105">This article presupposes that you are familiar with the basics of using the Office dialog API as described in [Use the Office dialog API in your Office Add-ins](dialog-api-in-office-add-ins.md).</span></span>
> 
> <span data-ttu-id="0297f-106">Voir aussi [meilleures pratiques et règles pour l’API de boîte de dialogue Office](dialog-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="0297f-106">See also [Best practices and rules for the Office dialog API](dialog-best-practices.md).</span></span>

<span data-ttu-id="0297f-107">Votre code doit gérer deux catégories d’événements :</span><span class="sxs-lookup"><span data-stu-id="0297f-107">Your code should handle two categories of events:</span></span>

- <span data-ttu-id="0297f-108">les erreurs renvoyées par l’appel de `displayDialogAsync` car la boîte de dialogue ne peut pas être créée ;</span><span class="sxs-lookup"><span data-stu-id="0297f-108">Errors returned by the call of `displayDialogAsync` because the dialog box cannot be created.</span></span>
- <span data-ttu-id="0297f-109">Erreurs et autres événements, dans la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="0297f-109">Errors, and other events, in the dialog box.</span></span>

## <a name="errors-from-displaydialogasync"></a><span data-ttu-id="0297f-110">Erreurs provenant de displayDialogAsync</span><span class="sxs-lookup"><span data-stu-id="0297f-110">Errors from displayDialogAsync</span></span>

<span data-ttu-id="0297f-111">En plus des erreurs système et de plateforme générales, quatre erreurs sont propres à l’appel `displayDialogAsync` .</span><span class="sxs-lookup"><span data-stu-id="0297f-111">In addition to general platform and system errors, four errors are specific to calling `displayDialogAsync`.</span></span>

|<span data-ttu-id="0297f-112">Numéro de code</span><span class="sxs-lookup"><span data-stu-id="0297f-112">Code number</span></span>|<span data-ttu-id="0297f-113">Signification</span><span class="sxs-lookup"><span data-stu-id="0297f-113">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="0297f-114">12004</span><span class="sxs-lookup"><span data-stu-id="0297f-114">12004</span></span>|<span data-ttu-id="0297f-p101">Le domaine de l’URL transmis à `displayDialogAsync` n’est pas approuvé. Le domaine doit être le même domaine que celui de la page hôte (y compris le protocole et le numéro de port).</span><span class="sxs-lookup"><span data-stu-id="0297f-p101">The domain of the URL passed to `displayDialogAsync` is not trusted. The domain must be the same domain as the host page (including protocol and port number).</span></span>|
|<span data-ttu-id="0297f-117">12005</span><span class="sxs-lookup"><span data-stu-id="0297f-117">12005</span></span>|<span data-ttu-id="0297f-118">L’URL transmise à `displayDialogAsync` utilise le protocole HTTP.</span><span class="sxs-lookup"><span data-stu-id="0297f-118">The URL passed to `displayDialogAsync` uses the HTTP protocol.</span></span> <span data-ttu-id="0297f-119">C’est le protocole HTTPS qui est requis.</span><span class="sxs-lookup"><span data-stu-id="0297f-119">HTTPS is required.</span></span> <span data-ttu-id="0297f-120">(Dans certaines versions d’Office, le texte du message d’erreur renvoyé avec 12005 est le même que celui renvoyé pour 12004.)</span><span class="sxs-lookup"><span data-stu-id="0297f-120">(In some versions of Office, the error message text returned with 12005 is the same one returned for 12004.)</span></span>|
|<span data-ttu-id="0297f-121"><span id="12007">12007</span></span><span class="sxs-lookup"><span data-stu-id="0297f-121"><span id="12007">12007</span></span></span><!-- The span is needed because office-js-helpers has an error message that links to this table row. -->|<span data-ttu-id="0297f-p103">Une boîte de dialogue est déjà ouverte à partir de cette fenêtre hôte. Une fenêtre hôte, par exemple un volet Office, ne peut avoir qu’une seule boîte de dialogue ouverte à la fois.</span><span class="sxs-lookup"><span data-stu-id="0297f-p103">A dialog box is already opened from this host window. A host window, such as a task pane, can only have one dialog box open at a time.</span></span>|
|<span data-ttu-id="0297f-124">12009</span><span class="sxs-lookup"><span data-stu-id="0297f-124">12009</span></span>|<span data-ttu-id="0297f-125">L’utilisateur a choisi d’ignorer la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="0297f-125">The user chose to ignore the dialog box.</span></span> <span data-ttu-id="0297f-126">Cette erreur peut se produire dans Office sur le Web, où les utilisateurs peuvent choisir de ne pas autoriser un complément à présenter une boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="0297f-126">This error can occur in Office on the web, where users may choose not to allow an add-in to present a dialog box.</span></span> <span data-ttu-id="0297f-127">Pour plus d’informations, consultez [la rubrique gestion des bloqueurs de fenêtres publicitaires intempestives avec Office sur le Web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="0297f-127">For more information, see [Handling pop-up blockers with Office on the web](dialog-best-practices.md#handling-pop-up-blockers-with-office-on-the-web).</span></span>|

<span data-ttu-id="0297f-128">Lorsque `displayDialogAsync` est appelé, il transmet un objet [asyncResult](/javascript/api/office/office.asyncresult) à sa fonction de rappel.</span><span class="sxs-lookup"><span data-stu-id="0297f-128">When `displayDialogAsync` is called, it passes an [AsyncResult](/javascript/api/office/office.asyncresult) object to its callback function.</span></span> <span data-ttu-id="0297f-129">Une fois l’appel réussi, la boîte de dialogue est ouverte et la `value` propriété de l' `AsyncResult` objet est un objet [Dialog](/javascript/api/office/office.dialog) .</span><span class="sxs-lookup"><span data-stu-id="0297f-129">When the call is successful, the dialog box is opened, and the `value` property of the `AsyncResult` object is a [Dialog](/javascript/api/office/office.dialog) object.</span></span> <span data-ttu-id="0297f-130">Pour obtenir un exemple, reportez-vous [à la rubrique envoyer des informations de la boîte de dialogue à la page hôte](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span><span class="sxs-lookup"><span data-stu-id="0297f-130">For an example of this, see [Send information from the dialog box to the host page](dialog-api-in-office-add-ins.md#send-information-from-the-dialog-box-to-the-host-page).</span></span> <span data-ttu-id="0297f-131">Lorsque l’appel à `displayDialogAsync` échoue, la boîte de dialogue n’est pas créée, la `status` propriété de l' `AsyncResult` objet est définie sur `Office.AsyncResultStatus.Failed` et la `error` propriété de l’objet est remplie.</span><span class="sxs-lookup"><span data-stu-id="0297f-131">When the call to `displayDialogAsync` fails, the dialog box is not created, the `status` property of the `AsyncResult` object is set to `Office.AsyncResultStatus.Failed`, and the `error` property of the object is populated.</span></span> <span data-ttu-id="0297f-132">Vous devez toujours fournir un rappel qui teste le `status` et répond lorsqu’il s’agit d’une erreur.</span><span class="sxs-lookup"><span data-stu-id="0297f-132">You should always provide a callback that tests the `status` and responds when it's an error.</span></span> <span data-ttu-id="0297f-133">Pour obtenir un exemple qui signale le message d’erreur quel que soit son numéro de code, consultez le code suivant.</span><span class="sxs-lookup"><span data-stu-id="0297f-133">For an example that reports the error message regardless of its code number, see the following code.</span></span> <span data-ttu-id="0297f-134">(La `showNotification` fonction, non définie dans cet article, affiche ou consigne l’erreur.</span><span class="sxs-lookup"><span data-stu-id="0297f-134">(The `showNotification` function, not defined in this article, either displays or logs the error.</span></span> <span data-ttu-id="0297f-135">Pour obtenir un exemple de la façon dont vous pouvez implémenter cette fonction dans votre complément, consultez la rubrique [exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span><span class="sxs-lookup"><span data-stu-id="0297f-135">For an example of how you can implement this function within your add-in, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).)</span></span>

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

## <a name="errors-and-events-in-the-dialog-box"></a><span data-ttu-id="0297f-136">Erreurs et événements dans la boîte de dialogue</span><span class="sxs-lookup"><span data-stu-id="0297f-136">Errors and events in the dialog box</span></span>

<span data-ttu-id="0297f-137">Trois erreurs et événements dans la boîte de dialogue déclencheront un `DialogEventReceived` événement dans la page hôte.</span><span class="sxs-lookup"><span data-stu-id="0297f-137">Three errors and events in the dialog box will raise a `DialogEventReceived` event in the host page.</span></span> <span data-ttu-id="0297f-138">Pour un rappel de ce qu’est une page hôte, voir [ouvrir une boîte de dialogue à partir d’une page hôte](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span><span class="sxs-lookup"><span data-stu-id="0297f-138">For a reminder of what a host page is, see [Open a dialog box from a host page](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page).</span></span>

|<span data-ttu-id="0297f-139">Numéro de code</span><span class="sxs-lookup"><span data-stu-id="0297f-139">Code number</span></span>|<span data-ttu-id="0297f-140">Signification</span><span class="sxs-lookup"><span data-stu-id="0297f-140">Meaning</span></span>|
|:-----|:-----|
|<span data-ttu-id="0297f-141">12002</span><span class="sxs-lookup"><span data-stu-id="0297f-141">12002</span></span>|<span data-ttu-id="0297f-142">Un des éléments suivants :</span><span class="sxs-lookup"><span data-stu-id="0297f-142">One of the following:</span></span><br> <span data-ttu-id="0297f-143">- Aucune page n’existe à l’URL qui a été transmise à `displayDialogAsync`.</span><span class="sxs-lookup"><span data-stu-id="0297f-143">- No page exists at the URL that was passed to `displayDialogAsync`.</span></span><br> <span data-ttu-id="0297f-144">-La page qui a été transmise au `displayDialogAsync` chargement, mais la boîte de dialogue a été redirigée vers une page qu’elle ne peut pas trouver ou chargée, ou elle a été dirigée vers une URL avec une syntaxe incorrecte.</span><span class="sxs-lookup"><span data-stu-id="0297f-144">- The page that was passed to `displayDialogAsync` loaded, but the dialog box was then redirected to a page that it cannot find or load, or it has been directed to a URL with invalid syntax.</span></span>|
|<span data-ttu-id="0297f-145">12003</span><span class="sxs-lookup"><span data-stu-id="0297f-145">12003</span></span>|<span data-ttu-id="0297f-p107">La boîte de dialogue a été redirigée vers une URL avec le protocole HTTP. C’est le protocole HTTPS qui est requis.</span><span class="sxs-lookup"><span data-stu-id="0297f-p107">The dialog box was directed to a URL with the HTTP protocol. HTTPS is required.</span></span>|
|<span data-ttu-id="0297f-148">12006</span><span class="sxs-lookup"><span data-stu-id="0297f-148">12006</span></span>|<span data-ttu-id="0297f-149">La boîte de dialogue a été fermée, généralement parce que l’utilisateur a cliqué sur le bouton **Fermer** **X**.</span><span class="sxs-lookup"><span data-stu-id="0297f-149">The dialog box was closed, usually because the user chose the **Close** button **X**.</span></span>|

<span data-ttu-id="0297f-p108">Votre code peut attribuer un gestionnaire pour l’événement `DialogEventReceived` dans l’appel de `displayDialogAsync`. Voici un exemple simple :</span><span class="sxs-lookup"><span data-stu-id="0297f-p108">Your code can assign a handler for the `DialogEventReceived` event in the call to `displayDialogAsync`. The following is a simple example:</span></span>

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Office.EventType.DialogEventReceived, processDialogEvent);
    }
);
```

<span data-ttu-id="0297f-152">Pour voir un exemple de gestionnaire pour l’événement `DialogEventReceived` qui crée des messages d’erreur personnalisés pour chaque code d’erreur, consultez l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="0297f-152">For an example of a handler for the `DialogEventReceived` event that creates custom error messages for each error code, see the following example:</span></span>

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

<span data-ttu-id="0297f-153">Pour voir un exemple de complément qui gère les erreurs de cette façon, consultez la rubrique relative à l’[exemple d’API de dialogue de complément Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span><span class="sxs-lookup"><span data-stu-id="0297f-153">For a sample add-in that handles errors in this way, see [Office Add-in Dialog API Example](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).</span></span>
