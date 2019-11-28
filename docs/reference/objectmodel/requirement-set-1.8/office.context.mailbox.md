---
title: Office. Context. Mailbox-ensemble de conditions requises 1,8
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 908eff7b34e63b62fbe250f1a6f810be69b17627
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629215"
---
# <a name="mailbox"></a><span data-ttu-id="63ab6-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="63ab6-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="63ab6-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="63ab6-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="63ab6-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="63ab6-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-105">Requirements</span></span>

|<span data-ttu-id="63ab6-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-106">Requirement</span></span>| <span data-ttu-id="63ab6-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-109">1.0</span></span>|
|[<span data-ttu-id="63ab6-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="63ab6-111">Restricted</span></span>|
|[<span data-ttu-id="63ab6-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="63ab6-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="63ab6-114">Members and methods</span></span>

| <span data-ttu-id="63ab6-115">Membre</span><span class="sxs-lookup"><span data-stu-id="63ab6-115">Member</span></span> | <span data-ttu-id="63ab6-116">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="63ab6-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="63ab6-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="63ab6-118">Membre</span><span class="sxs-lookup"><span data-stu-id="63ab6-118">Member</span></span> |
| [<span data-ttu-id="63ab6-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="63ab6-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="63ab6-120">Membre</span><span class="sxs-lookup"><span data-stu-id="63ab6-120">Member</span></span> |
| [<span data-ttu-id="63ab6-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="63ab6-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="63ab6-122">Membre</span><span class="sxs-lookup"><span data-stu-id="63ab6-122">Member</span></span> |
| [<span data-ttu-id="63ab6-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="63ab6-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="63ab6-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-124">Method</span></span> |
| [<span data-ttu-id="63ab6-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="63ab6-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="63ab6-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-126">Method</span></span> |
| [<span data-ttu-id="63ab6-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="63ab6-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="63ab6-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-128">Method</span></span> |
| [<span data-ttu-id="63ab6-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="63ab6-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="63ab6-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-130">Method</span></span> |
| [<span data-ttu-id="63ab6-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="63ab6-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="63ab6-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-132">Method</span></span> |
| [<span data-ttu-id="63ab6-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="63ab6-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="63ab6-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-134">Method</span></span> |
| [<span data-ttu-id="63ab6-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="63ab6-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="63ab6-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-136">Method</span></span> |
| [<span data-ttu-id="63ab6-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="63ab6-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="63ab6-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-138">Method</span></span> |
| [<span data-ttu-id="63ab6-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="63ab6-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="63ab6-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-140">Method</span></span> |
| [<span data-ttu-id="63ab6-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="63ab6-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="63ab6-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-142">Method</span></span> |
| [<span data-ttu-id="63ab6-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="63ab6-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="63ab6-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-144">Method</span></span> |
| [<span data-ttu-id="63ab6-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="63ab6-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="63ab6-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-146">Method</span></span> |
| [<span data-ttu-id="63ab6-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="63ab6-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="63ab6-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-148">Method</span></span> |
| [<span data-ttu-id="63ab6-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="63ab6-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="63ab6-150">Méthode</span><span class="sxs-lookup"><span data-stu-id="63ab6-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="63ab6-151">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="63ab6-151">Namespaces</span></span>

<span data-ttu-id="63ab6-152">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="63ab6-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="63ab6-153">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="63ab6-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="63ab6-154">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="63ab6-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="63ab6-155">Members</span><span class="sxs-lookup"><span data-stu-id="63ab6-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="63ab6-156">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="63ab6-156">ewsUrl: String</span></span>

<span data-ttu-id="63ab6-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-159">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="63ab6-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="63ab6-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="63ab6-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="63ab6-162">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="63ab6-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="63ab6-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="63ab6-165">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-165">Type</span></span>

*   <span data-ttu-id="63ab6-166">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="63ab6-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-167">Requirements</span></span>

|<span data-ttu-id="63ab6-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-168">Requirement</span></span>| <span data-ttu-id="63ab6-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-171">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-171">1.0</span></span>|
|[<span data-ttu-id="63ab6-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-173">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategoriesviewoutlook-js-18"></a><span data-ttu-id="63ab6-176">masterCategories : [masterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="63ab6-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="63ab6-177">Obtient un objet qui fournit des méthodes pour gérer la liste de formes de base des catégories sur cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="63ab6-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-178">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="63ab6-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="63ab6-179">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-179">Type</span></span>

*   [<span data-ttu-id="63ab6-180">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="63ab6-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="63ab6-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-181">Requirements</span></span>

|<span data-ttu-id="63ab6-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-182">Requirement</span></span>| <span data-ttu-id="63ab6-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-185">1.8</span><span class="sxs-lookup"><span data-stu-id="63ab6-185">1.8</span></span> |
|[<span data-ttu-id="63ab6-186">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="63ab6-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="63ab6-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="63ab6-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-190">Example</span></span>

<span data-ttu-id="63ab6-191">Cet exemple obtient la liste principale des catégories pour cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="63ab6-191">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="63ab6-192">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="63ab6-192">restUrl: String</span></span>

<span data-ttu-id="63ab6-193">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="63ab6-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="63ab6-194">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="63ab6-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="63ab6-195">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-195">Type</span></span>

*   <span data-ttu-id="63ab6-196">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-196">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="63ab6-197">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-197">Requirements</span></span>

|<span data-ttu-id="63ab6-198">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-198">Requirement</span></span>| <span data-ttu-id="63ab6-199">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-200">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-201">1,5</span><span class="sxs-lookup"><span data-stu-id="63ab6-201">1.5</span></span> |
|[<span data-ttu-id="63ab6-202">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-203">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-204">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-205">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-205">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="63ab6-206">Méthodes</span><span class="sxs-lookup"><span data-stu-id="63ab6-206">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="63ab6-207">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="63ab6-207">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="63ab6-208">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="63ab6-208">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="63ab6-209">Actuellement, les types d’événement pris `Office.EventType.ItemChanged` en `Office.EventType.OfficeThemeChanged`charge sont et.</span><span class="sxs-lookup"><span data-stu-id="63ab6-209">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-210">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-210">Parameters</span></span>

| <span data-ttu-id="63ab6-211">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-211">Name</span></span> | <span data-ttu-id="63ab6-212">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-212">Type</span></span> | <span data-ttu-id="63ab6-213">Attributs</span><span class="sxs-lookup"><span data-stu-id="63ab6-213">Attributes</span></span> | <span data-ttu-id="63ab6-214">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-214">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="63ab6-215">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="63ab6-215">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="63ab6-216">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="63ab6-216">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="63ab6-217">Fonction</span><span class="sxs-lookup"><span data-stu-id="63ab6-217">Function</span></span> || <span data-ttu-id="63ab6-p104">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="63ab6-221">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-221">Object</span></span> | <span data-ttu-id="63ab6-222">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-222">&lt;optional&gt;</span></span> | <span data-ttu-id="63ab6-223">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="63ab6-223">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="63ab6-224">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-224">Object</span></span> | <span data-ttu-id="63ab6-225">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-225">&lt;optional&gt;</span></span> | <span data-ttu-id="63ab6-226">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="63ab6-226">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="63ab6-227">fonction</span><span class="sxs-lookup"><span data-stu-id="63ab6-227">function</span></span>| <span data-ttu-id="63ab6-228">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-228">&lt;optional&gt;</span></span>|<span data-ttu-id="63ab6-229">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="63ab6-229">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-230">Requirements</span></span>

|<span data-ttu-id="63ab6-231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-231">Requirement</span></span>| <span data-ttu-id="63ab6-232">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-234">1,5</span><span class="sxs-lookup"><span data-stu-id="63ab6-234">1.5</span></span> |
|[<span data-ttu-id="63ab6-235">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-236">ReadItem</span></span> |
|[<span data-ttu-id="63ab6-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-239">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="63ab6-240">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="63ab6-240">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="63ab6-241">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="63ab6-241">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-242">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="63ab6-242">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="63ab6-p105">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-245">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-245">Parameters</span></span>

|<span data-ttu-id="63ab6-246">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-246">Name</span></span>| <span data-ttu-id="63ab6-247">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-247">Type</span></span>| <span data-ttu-id="63ab6-248">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="63ab6-249">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-249">String</span></span>|<span data-ttu-id="63ab6-250">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="63ab6-250">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="63ab6-251">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="63ab6-251">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="63ab6-252">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-252">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-253">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-253">Requirements</span></span>

|<span data-ttu-id="63ab6-254">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-254">Requirement</span></span>| <span data-ttu-id="63ab6-255">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-256">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-257">1.3</span><span class="sxs-lookup"><span data-stu-id="63ab6-257">1.3</span></span>|
|[<span data-ttu-id="63ab6-258">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-258">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-259">Restreinte</span><span class="sxs-lookup"><span data-stu-id="63ab6-259">Restricted</span></span>|
|[<span data-ttu-id="63ab6-260">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-260">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-261">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-261">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="63ab6-262">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="63ab6-262">Returns:</span></span>

<span data-ttu-id="63ab6-263">Type : String</span><span class="sxs-lookup"><span data-stu-id="63ab6-263">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="63ab6-264">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-264">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-18"></a><span data-ttu-id="63ab6-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="63ab6-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="63ab6-266">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="63ab6-266">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="63ab6-p106">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="63ab6-p107">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-272">Paramètres</span><span class="sxs-lookup"><span data-stu-id="63ab6-272">Parameters</span></span>

|<span data-ttu-id="63ab6-273">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-273">Name</span></span>| <span data-ttu-id="63ab6-274">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-274">Type</span></span>| <span data-ttu-id="63ab6-275">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-275">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="63ab6-276">Date</span><span class="sxs-lookup"><span data-stu-id="63ab6-276">Date</span></span>|<span data-ttu-id="63ab6-277">Objet Date</span><span class="sxs-lookup"><span data-stu-id="63ab6-277">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-278">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-278">Requirements</span></span>

|<span data-ttu-id="63ab6-279">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-279">Requirement</span></span>| <span data-ttu-id="63ab6-280">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-281">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-281">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-282">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-282">1.0</span></span>|
|[<span data-ttu-id="63ab6-283">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-283">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-284">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-285">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-285">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-286">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-286">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="63ab6-287">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="63ab6-287">Returns:</span></span>

<span data-ttu-id="63ab6-288">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="63ab6-288">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="63ab6-289">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="63ab6-289">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="63ab6-290">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="63ab6-290">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-291">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="63ab6-291">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="63ab6-p108">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-294">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-294">Parameters</span></span>

|<span data-ttu-id="63ab6-295">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-295">Name</span></span>| <span data-ttu-id="63ab6-296">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-296">Type</span></span>| <span data-ttu-id="63ab6-297">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-297">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="63ab6-298">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-298">String</span></span>|<span data-ttu-id="63ab6-299">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="63ab6-299">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="63ab6-300">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="63ab6-300">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="63ab6-301">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="63ab6-301">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-302">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-302">Requirements</span></span>

|<span data-ttu-id="63ab6-303">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-303">Requirement</span></span>| <span data-ttu-id="63ab6-304">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-305">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-306">1.3</span><span class="sxs-lookup"><span data-stu-id="63ab6-306">1.3</span></span>|
|[<span data-ttu-id="63ab6-307">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-308">Restreinte</span><span class="sxs-lookup"><span data-stu-id="63ab6-308">Restricted</span></span>|
|[<span data-ttu-id="63ab6-309">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-310">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-310">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="63ab6-311">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="63ab6-311">Returns:</span></span>

<span data-ttu-id="63ab6-312">Type : String</span><span class="sxs-lookup"><span data-stu-id="63ab6-312">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="63ab6-313">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-313">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="63ab6-314">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="63ab6-314">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="63ab6-315">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="63ab6-315">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="63ab6-316">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="63ab6-316">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-317">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-317">Parameters</span></span>

|<span data-ttu-id="63ab6-318">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-318">Name</span></span>| <span data-ttu-id="63ab6-319">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-319">Type</span></span>| <span data-ttu-id="63ab6-320">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-320">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="63ab6-321">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="63ab6-321">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)|<span data-ttu-id="63ab6-322">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="63ab6-322">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-323">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-323">Requirements</span></span>

|<span data-ttu-id="63ab6-324">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-324">Requirement</span></span>| <span data-ttu-id="63ab6-325">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-326">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-327">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-327">1.0</span></span>|
|[<span data-ttu-id="63ab6-328">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-329">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-330">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-331">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-331">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="63ab6-332">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="63ab6-332">Returns:</span></span>

<span data-ttu-id="63ab6-333">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="63ab6-333">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="63ab6-334">Type : Date</span><span class="sxs-lookup"><span data-stu-id="63ab6-334">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="63ab6-335">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-335">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="63ab6-336">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="63ab6-336">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="63ab6-337">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="63ab6-337">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-338">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="63ab6-338">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="63ab6-339">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="63ab6-339">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="63ab6-p109">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="63ab6-342">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-342">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="63ab6-343">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="63ab6-343">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-344">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-344">Parameters</span></span>

|<span data-ttu-id="63ab6-345">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-345">Name</span></span>| <span data-ttu-id="63ab6-346">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-346">Type</span></span>| <span data-ttu-id="63ab6-347">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-347">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="63ab6-348">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-348">String</span></span>|<span data-ttu-id="63ab6-349">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="63ab6-349">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-350">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-350">Requirements</span></span>

|<span data-ttu-id="63ab6-351">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-351">Requirement</span></span>| <span data-ttu-id="63ab6-352">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-353">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-354">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-354">1.0</span></span>|
|[<span data-ttu-id="63ab6-355">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-356">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-357">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-358">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-358">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-359">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-359">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="63ab6-360">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="63ab6-360">displayMessageForm(itemId)</span></span>

<span data-ttu-id="63ab6-361">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="63ab6-361">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-362">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="63ab6-362">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="63ab6-363">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="63ab6-363">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="63ab6-364">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-364">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="63ab6-365">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="63ab6-365">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="63ab6-p110">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-368">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-368">Parameters</span></span>

|<span data-ttu-id="63ab6-369">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-369">Name</span></span>| <span data-ttu-id="63ab6-370">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-370">Type</span></span>| <span data-ttu-id="63ab6-371">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-371">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="63ab6-372">Chaîne</span><span class="sxs-lookup"><span data-stu-id="63ab6-372">String</span></span>|<span data-ttu-id="63ab6-373">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="63ab6-373">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-374">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-374">Requirements</span></span>

|<span data-ttu-id="63ab6-375">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-375">Requirement</span></span>| <span data-ttu-id="63ab6-376">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-377">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-378">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-378">1.0</span></span>|
|[<span data-ttu-id="63ab6-379">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-380">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-381">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-382">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-382">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-383">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-383">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="63ab6-384">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="63ab6-384">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="63ab6-385">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="63ab6-385">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-386">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="63ab6-386">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="63ab6-p111">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="63ab6-p112">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="63ab6-p113">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="63ab6-394">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="63ab6-394">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-395">Paramètres</span><span class="sxs-lookup"><span data-stu-id="63ab6-395">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-396">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="63ab6-396">All parameters are optional.</span></span>

|<span data-ttu-id="63ab6-397">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-397">Name</span></span>| <span data-ttu-id="63ab6-398">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-398">Type</span></span>| <span data-ttu-id="63ab6-399">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-399">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="63ab6-400">Object</span><span class="sxs-lookup"><span data-stu-id="63ab6-400">Object</span></span> | <span data-ttu-id="63ab6-401">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="63ab6-401">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="63ab6-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="63ab6-p114">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="63ab6-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="63ab6-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="63ab6-408">Date</span><span class="sxs-lookup"><span data-stu-id="63ab6-408">Date</span></span> | <span data-ttu-id="63ab6-409">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="63ab6-409">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="63ab6-410">Date</span><span class="sxs-lookup"><span data-stu-id="63ab6-410">Date</span></span> | <span data-ttu-id="63ab6-411">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="63ab6-411">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="63ab6-412">Chaîne</span><span class="sxs-lookup"><span data-stu-id="63ab6-412">String</span></span> | <span data-ttu-id="63ab6-p116">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="63ab6-415">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-415">Array.&lt;String&gt;</span></span> | <span data-ttu-id="63ab6-p117">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="63ab6-418">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-418">String</span></span> | <span data-ttu-id="63ab6-p118">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="63ab6-421">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-421">String</span></span> | <span data-ttu-id="63ab6-p119">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="63ab6-424">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-424">Requirements</span></span>

|<span data-ttu-id="63ab6-425">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-425">Requirement</span></span>| <span data-ttu-id="63ab6-426">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-427">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-428">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-428">1.0</span></span>|
|[<span data-ttu-id="63ab6-429">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-430">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-431">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-432">Lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-432">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-433">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-433">Example</span></span>

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="63ab6-434">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="63ab6-434">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="63ab6-435">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="63ab6-435">Displays a form for creating a new message.</span></span>

<span data-ttu-id="63ab6-436">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="63ab6-436">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="63ab6-437">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="63ab6-437">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="63ab6-438">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="63ab6-438">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-439">Paramètres</span><span class="sxs-lookup"><span data-stu-id="63ab6-439">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-440">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="63ab6-440">All parameters are optional.</span></span>

|<span data-ttu-id="63ab6-441">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-441">Name</span></span>| <span data-ttu-id="63ab6-442">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-442">Type</span></span>| <span data-ttu-id="63ab6-443">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-443">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="63ab6-444">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-444">Object</span></span> | <span data-ttu-id="63ab6-445">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="63ab6-445">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="63ab6-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="63ab6-447">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="63ab6-447">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="63ab6-448">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-448">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="63ab6-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="63ab6-450">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="63ab6-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="63ab6-451">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="63ab6-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="63ab6-453">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="63ab6-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="63ab6-454">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="63ab6-455">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-455">String</span></span> | <span data-ttu-id="63ab6-456">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="63ab6-456">A string containing the subject of the message.</span></span> <span data-ttu-id="63ab6-457">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="63ab6-457">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="63ab6-458">Chaîne</span><span class="sxs-lookup"><span data-stu-id="63ab6-458">String</span></span> | <span data-ttu-id="63ab6-459">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="63ab6-459">The HTML body of the message.</span></span> <span data-ttu-id="63ab6-460">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="63ab6-460">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="63ab6-461">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-461">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="63ab6-462">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-462">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="63ab6-463">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-463">String</span></span> | <span data-ttu-id="63ab6-p126">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="63ab6-466">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-466">String</span></span> | <span data-ttu-id="63ab6-467">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="63ab6-467">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="63ab6-468">Chaîne</span><span class="sxs-lookup"><span data-stu-id="63ab6-468">String</span></span> | <span data-ttu-id="63ab6-p127">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="63ab6-471">Booléen</span><span class="sxs-lookup"><span data-stu-id="63ab6-471">Boolean</span></span> | <span data-ttu-id="63ab6-p128">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="63ab6-474">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-474">String</span></span> | <span data-ttu-id="63ab6-475">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-475">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="63ab6-476">ID d’élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="63ab6-476">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="63ab6-477">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="63ab6-477">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="63ab6-478">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-478">Requirements</span></span>

|<span data-ttu-id="63ab6-479">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-479">Requirement</span></span>| <span data-ttu-id="63ab6-480">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-481">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-482">1.6</span><span class="sxs-lookup"><span data-stu-id="63ab6-482">1.6</span></span> |
|[<span data-ttu-id="63ab6-483">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-484">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-485">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-486">Lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-487">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-487">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="63ab6-488">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="63ab6-488">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="63ab6-489">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="63ab6-489">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="63ab6-p130">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-492">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="63ab6-492">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="63ab6-493">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="63ab6-493">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="63ab6-494">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-494">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="63ab6-495">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="63ab6-495">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="63ab6-496">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="63ab6-496">**REST Tokens**</span></span>

<span data-ttu-id="63ab6-p132">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="63ab6-500">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="63ab6-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="63ab6-501">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="63ab6-501">**EWS Tokens**</span></span>

<span data-ttu-id="63ab6-p133">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="63ab6-504">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="63ab6-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="63ab6-505">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="63ab6-505">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="63ab6-506">Le système tiers utilise le jeton comme jeton d’autorisation du support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) des services Web Exchange (EWS) ou de [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour récupérer une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-506">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="63ab6-507">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="63ab6-507">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-508">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-508">Parameters</span></span>

|<span data-ttu-id="63ab6-509">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-509">Name</span></span>| <span data-ttu-id="63ab6-510">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-510">Type</span></span>| <span data-ttu-id="63ab6-511">Attributs</span><span class="sxs-lookup"><span data-stu-id="63ab6-511">Attributes</span></span>| <span data-ttu-id="63ab6-512">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-512">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="63ab6-513">Object</span><span class="sxs-lookup"><span data-stu-id="63ab6-513">Object</span></span> | <span data-ttu-id="63ab6-514">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-514">&lt;optional&gt;</span></span> | <span data-ttu-id="63ab6-515">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="63ab6-515">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="63ab6-516">Boolean</span><span class="sxs-lookup"><span data-stu-id="63ab6-516">Boolean</span></span> |  <span data-ttu-id="63ab6-517">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-517">&lt;optional&gt;</span></span> | <span data-ttu-id="63ab6-p135">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="63ab6-520">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-520">Object</span></span> |  <span data-ttu-id="63ab6-521">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-521">&lt;optional&gt;</span></span> | <span data-ttu-id="63ab6-522">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="63ab6-522">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="63ab6-523">fonction</span><span class="sxs-lookup"><span data-stu-id="63ab6-523">function</span></span>||<span data-ttu-id="63ab6-524">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="63ab6-524">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="63ab6-525">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-525">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="63ab6-526">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="63ab6-526">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="63ab6-527">Erreurs</span><span class="sxs-lookup"><span data-stu-id="63ab6-527">Errors</span></span>

|<span data-ttu-id="63ab6-528">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="63ab6-528">Error code</span></span>|<span data-ttu-id="63ab6-529">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-529">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="63ab6-530">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="63ab6-530">The request has failed.</span></span> <span data-ttu-id="63ab6-531">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="63ab6-531">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="63ab6-532">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="63ab6-532">The Exchange server returned an error.</span></span> <span data-ttu-id="63ab6-533">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="63ab6-533">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="63ab6-534">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="63ab6-534">The user is no longer connected to the network.</span></span> <span data-ttu-id="63ab6-535">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="63ab6-535">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-536">Requirements</span></span>

|<span data-ttu-id="63ab6-537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-537">Requirement</span></span>| <span data-ttu-id="63ab6-538">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-539">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-540">1,5</span><span class="sxs-lookup"><span data-stu-id="63ab6-540">1.5</span></span> |
|[<span data-ttu-id="63ab6-541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-542">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-543">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-544">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-545">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-545">Example</span></span>

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="63ab6-546">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="63ab6-546">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="63ab6-547">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="63ab6-547">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="63ab6-p139">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="63ab6-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="63ab6-550">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="63ab6-550">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="63ab6-551">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-551">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="63ab6-552">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="63ab6-552">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="63ab6-553">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="63ab6-553">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="63ab6-554">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-554">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="63ab6-555">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="63ab6-555">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-556">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-556">Parameters</span></span>

|<span data-ttu-id="63ab6-557">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-557">Name</span></span>| <span data-ttu-id="63ab6-558">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-558">Type</span></span>| <span data-ttu-id="63ab6-559">Attributs</span><span class="sxs-lookup"><span data-stu-id="63ab6-559">Attributes</span></span>| <span data-ttu-id="63ab6-560">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-560">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="63ab6-561">function</span><span class="sxs-lookup"><span data-stu-id="63ab6-561">function</span></span>||<span data-ttu-id="63ab6-562">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="63ab6-562">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="63ab6-563">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-563">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="63ab6-564">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="63ab6-564">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="63ab6-565">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-565">Object</span></span>| <span data-ttu-id="63ab6-566">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-566">&lt;optional&gt;</span></span>|<span data-ttu-id="63ab6-567">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="63ab6-567">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="63ab6-568">Erreurs</span><span class="sxs-lookup"><span data-stu-id="63ab6-568">Errors</span></span>

|<span data-ttu-id="63ab6-569">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="63ab6-569">Error code</span></span>|<span data-ttu-id="63ab6-570">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-570">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="63ab6-571">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="63ab6-571">The request has failed.</span></span> <span data-ttu-id="63ab6-572">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="63ab6-572">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="63ab6-573">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="63ab6-573">The Exchange server returned an error.</span></span> <span data-ttu-id="63ab6-574">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="63ab6-574">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="63ab6-575">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="63ab6-575">The user is no longer connected to the network.</span></span> <span data-ttu-id="63ab6-576">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="63ab6-576">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-577">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-577">Requirements</span></span>

|<span data-ttu-id="63ab6-578">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-578">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="63ab6-579">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-579">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-580">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-580">1.0</span></span> | <span data-ttu-id="63ab6-581">1.3</span><span class="sxs-lookup"><span data-stu-id="63ab6-581">1.3</span></span> |
|[<span data-ttu-id="63ab6-582">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-583">ReadItem</span></span> | <span data-ttu-id="63ab6-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-584">ReadItem</span></span> |
|[<span data-ttu-id="63ab6-585">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-586">Lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-586">Read</span></span> | <span data-ttu-id="63ab6-587">Composition</span><span class="sxs-lookup"><span data-stu-id="63ab6-587">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="63ab6-588">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-588">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="63ab6-589">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="63ab6-589">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="63ab6-590">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="63ab6-590">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="63ab6-591">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="63ab6-591">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-592">Paramètres</span><span class="sxs-lookup"><span data-stu-id="63ab6-592">Parameters</span></span>

|<span data-ttu-id="63ab6-593">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-593">Name</span></span>| <span data-ttu-id="63ab6-594">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-594">Type</span></span>| <span data-ttu-id="63ab6-595">Attributs</span><span class="sxs-lookup"><span data-stu-id="63ab6-595">Attributes</span></span>| <span data-ttu-id="63ab6-596">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-596">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="63ab6-597">fonction</span><span class="sxs-lookup"><span data-stu-id="63ab6-597">function</span></span>||<span data-ttu-id="63ab6-598">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="63ab6-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="63ab6-599">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-599">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="63ab6-600">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="63ab6-600">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="63ab6-601">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-601">Object</span></span>| <span data-ttu-id="63ab6-602">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-602">&lt;optional&gt;</span></span>|<span data-ttu-id="63ab6-603">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="63ab6-603">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="63ab6-604">Erreurs</span><span class="sxs-lookup"><span data-stu-id="63ab6-604">Errors</span></span>

|<span data-ttu-id="63ab6-605">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="63ab6-605">Error code</span></span>|<span data-ttu-id="63ab6-606">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-606">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="63ab6-607">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="63ab6-607">The request has failed.</span></span> <span data-ttu-id="63ab6-608">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="63ab6-608">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="63ab6-609">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="63ab6-609">The Exchange server returned an error.</span></span> <span data-ttu-id="63ab6-610">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="63ab6-610">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="63ab6-611">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="63ab6-611">The user is no longer connected to the network.</span></span> <span data-ttu-id="63ab6-612">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="63ab6-612">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-613">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-613">Requirements</span></span>

|<span data-ttu-id="63ab6-614">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-614">Requirement</span></span>| <span data-ttu-id="63ab6-615">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-616">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-617">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-617">1.0</span></span>|
|[<span data-ttu-id="63ab6-618">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-619">ReadItem</span></span>|
|[<span data-ttu-id="63ab6-620">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-621">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-622">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-622">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="63ab6-623">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="63ab6-623">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="63ab6-624">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="63ab6-624">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-625">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="63ab6-625">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="63ab6-626">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="63ab6-626">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="63ab6-627">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="63ab6-627">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="63ab6-628">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="63ab6-628">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="63ab6-629">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-629">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="63ab6-630">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="63ab6-630">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="63ab6-631">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-631">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="63ab6-632">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="63ab6-632">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="63ab6-p149">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="63ab6-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="63ab6-635">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="63ab6-635">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="63ab6-636">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="63ab6-636">Version differences</span></span>

<span data-ttu-id="63ab6-637">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-637">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="63ab6-638">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage.</span><span class="sxs-lookup"><span data-stu-id="63ab6-638">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="63ab6-639">Vous pouvez déterminer si votre application de messagerie est en cours d’exécution dans Outlook sur le Web ou sur un client de bureau à l’aide de la propriété Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="63ab6-639">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="63ab6-640">Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="63ab6-640">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-641">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-641">Parameters</span></span>

|<span data-ttu-id="63ab6-642">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-642">Name</span></span>| <span data-ttu-id="63ab6-643">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-643">Type</span></span>| <span data-ttu-id="63ab6-644">Attributs</span><span class="sxs-lookup"><span data-stu-id="63ab6-644">Attributes</span></span>| <span data-ttu-id="63ab6-645">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-645">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="63ab6-646">String</span><span class="sxs-lookup"><span data-stu-id="63ab6-646">String</span></span>||<span data-ttu-id="63ab6-647">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="63ab6-647">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="63ab6-648">function</span><span class="sxs-lookup"><span data-stu-id="63ab6-648">function</span></span>||<span data-ttu-id="63ab6-649">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="63ab6-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="63ab6-650">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="63ab6-650">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="63ab6-651">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="63ab6-651">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="63ab6-652">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-652">Object</span></span>| <span data-ttu-id="63ab6-653">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-653">&lt;optional&gt;</span></span>|<span data-ttu-id="63ab6-654">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="63ab6-654">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-655">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-655">Requirements</span></span>

|<span data-ttu-id="63ab6-656">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-656">Requirement</span></span>| <span data-ttu-id="63ab6-657">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-657">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-658">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-658">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-659">1.0</span><span class="sxs-lookup"><span data-stu-id="63ab6-659">1.0</span></span>|
|[<span data-ttu-id="63ab6-660">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-660">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-661">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="63ab6-661">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="63ab6-662">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-662">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-663">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-663">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="63ab6-664">Exemple</span><span class="sxs-lookup"><span data-stu-id="63ab6-664">Example</span></span>

<span data-ttu-id="63ab6-665">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="63ab6-665">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="63ab6-666">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="63ab6-666">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="63ab6-667">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="63ab6-667">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="63ab6-668">Actuellement, les types d’événement pris `Office.EventType.ItemChanged` en `Office.EventType.OfficeThemeChanged`charge sont et.</span><span class="sxs-lookup"><span data-stu-id="63ab6-668">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="63ab6-669">Parameters</span><span class="sxs-lookup"><span data-stu-id="63ab6-669">Parameters</span></span>

| <span data-ttu-id="63ab6-670">Nom</span><span class="sxs-lookup"><span data-stu-id="63ab6-670">Name</span></span> | <span data-ttu-id="63ab6-671">Type</span><span class="sxs-lookup"><span data-stu-id="63ab6-671">Type</span></span> | <span data-ttu-id="63ab6-672">Attributs</span><span class="sxs-lookup"><span data-stu-id="63ab6-672">Attributes</span></span> | <span data-ttu-id="63ab6-673">Description</span><span class="sxs-lookup"><span data-stu-id="63ab6-673">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="63ab6-674">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="63ab6-674">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="63ab6-675">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="63ab6-675">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="63ab6-676">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-676">Object</span></span> | <span data-ttu-id="63ab6-677">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-677">&lt;optional&gt;</span></span> | <span data-ttu-id="63ab6-678">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="63ab6-678">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="63ab6-679">Objet</span><span class="sxs-lookup"><span data-stu-id="63ab6-679">Object</span></span> | <span data-ttu-id="63ab6-680">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-680">&lt;optional&gt;</span></span> | <span data-ttu-id="63ab6-681">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="63ab6-681">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="63ab6-682">fonction</span><span class="sxs-lookup"><span data-stu-id="63ab6-682">function</span></span>| <span data-ttu-id="63ab6-683">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="63ab6-683">&lt;optional&gt;</span></span>|<span data-ttu-id="63ab6-684">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="63ab6-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="63ab6-685">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="63ab6-685">Requirements</span></span>

|<span data-ttu-id="63ab6-686">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="63ab6-686">Requirement</span></span>| <span data-ttu-id="63ab6-687">Valeur</span><span class="sxs-lookup"><span data-stu-id="63ab6-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="63ab6-688">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="63ab6-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="63ab6-689">1,5</span><span class="sxs-lookup"><span data-stu-id="63ab6-689">1.5</span></span> |
|[<span data-ttu-id="63ab6-690">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="63ab6-690">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="63ab6-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="63ab6-691">ReadItem</span></span> |
|[<span data-ttu-id="63ab6-692">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="63ab6-692">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="63ab6-693">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="63ab6-693">Compose or Read</span></span>|
