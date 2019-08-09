---
title: Office. Context. Mailbox-Preview-ensemble de conditions requises
description: ''
ms.date: 08/06/2019
localization_priority: Normal
ms.openlocfilehash: 8f7345532e3c65f9945f424520c37f707634bf3b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268718"
---
# <a name="mailbox"></a><span data-ttu-id="4846e-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="4846e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="4846e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="4846e-104">Permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="4846e-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4846e-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-105">Requirements</span></span>

|<span data-ttu-id="4846e-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-106">Requirement</span></span>| <span data-ttu-id="4846e-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-109">1.0</span></span>|
|[<span data-ttu-id="4846e-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4846e-111">Restricted</span></span>|
|[<span data-ttu-id="4846e-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4846e-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4846e-114">Members and methods</span></span>

| <span data-ttu-id="4846e-115">Membre</span><span class="sxs-lookup"><span data-stu-id="4846e-115">Member</span></span> | <span data-ttu-id="4846e-116">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4846e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="4846e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="4846e-118">Membre</span><span class="sxs-lookup"><span data-stu-id="4846e-118">Member</span></span> |
| [<span data-ttu-id="4846e-119">masterCategories</span><span class="sxs-lookup"><span data-stu-id="4846e-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="4846e-120">Membre</span><span class="sxs-lookup"><span data-stu-id="4846e-120">Member</span></span> |
| [<span data-ttu-id="4846e-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="4846e-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="4846e-122">Membre</span><span class="sxs-lookup"><span data-stu-id="4846e-122">Member</span></span> |
| [<span data-ttu-id="4846e-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4846e-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4846e-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-124">Method</span></span> |
| [<span data-ttu-id="4846e-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="4846e-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="4846e-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-126">Method</span></span> |
| [<span data-ttu-id="4846e-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4846e-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="4846e-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-128">Method</span></span> |
| [<span data-ttu-id="4846e-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="4846e-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="4846e-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-130">Method</span></span> |
| [<span data-ttu-id="4846e-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="4846e-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="4846e-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-132">Method</span></span> |
| [<span data-ttu-id="4846e-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4846e-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="4846e-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-134">Method</span></span> |
| [<span data-ttu-id="4846e-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="4846e-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="4846e-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-136">Method</span></span> |
| [<span data-ttu-id="4846e-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4846e-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="4846e-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-138">Method</span></span> |
| [<span data-ttu-id="4846e-139">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="4846e-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="4846e-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-140">Method</span></span> |
| [<span data-ttu-id="4846e-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4846e-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="4846e-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-142">Method</span></span> |
| [<span data-ttu-id="4846e-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4846e-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="4846e-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-144">Method</span></span> |
| [<span data-ttu-id="4846e-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4846e-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="4846e-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-146">Method</span></span> |
| [<span data-ttu-id="4846e-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="4846e-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="4846e-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-148">Method</span></span> |
| [<span data-ttu-id="4846e-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4846e-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4846e-150">Méthode</span><span class="sxs-lookup"><span data-stu-id="4846e-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4846e-151">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4846e-151">Namespaces</span></span>

<span data-ttu-id="4846e-152">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4846e-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4846e-153">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4846e-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4846e-154">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4846e-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4846e-155">Membres</span><span class="sxs-lookup"><span data-stu-id="4846e-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4846e-156">ewsUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-156">ewsUrl: String</span></span>

<span data-ttu-id="4846e-157">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4846e-157">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="4846e-158">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4846e-158">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-159">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="4846e-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4846e-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4846e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4846e-162">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="4846e-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="4846e-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4846e-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4846e-165">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-165">Type</span></span>

*   <span data-ttu-id="4846e-166">String</span><span class="sxs-lookup"><span data-stu-id="4846e-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4846e-167">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-167">Requirements</span></span>

|<span data-ttu-id="4846e-168">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-168">Requirement</span></span>| <span data-ttu-id="4846e-169">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-170">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-171">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-171">1.0</span></span>|
|[<span data-ttu-id="4846e-172">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-173">ReadItem</span></span>|
|[<span data-ttu-id="4846e-174">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-175">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-175">Compose or Read</span></span>|

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="4846e-176">masterCategories: [masterCategories](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="4846e-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="4846e-177">Obtient un objet qui fournit des méthodes pour gérer la liste de formes de base des catégories sur cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="4846e-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-178">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="4846e-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="4846e-179">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-179">Type</span></span>

*   [<span data-ttu-id="4846e-180">Catégoriesmaître</span><span class="sxs-lookup"><span data-stu-id="4846e-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="4846e-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-181">Requirements</span></span>

|<span data-ttu-id="4846e-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-182">Requirement</span></span>| <span data-ttu-id="4846e-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-185">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4846e-185">Preview</span></span> |
|[<span data-ttu-id="4846e-186">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4846e-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="4846e-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="4846e-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-190">Example</span></span>

<span data-ttu-id="4846e-191">Cet exemple obtient la liste principale des catégories pour cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="4846e-191">This example gets the categories master list for this mailbox.</span></span>

```javascript
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="resturl-string"></a><span data-ttu-id="4846e-192">restUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-192">restUrl: String</span></span>

<span data-ttu-id="4846e-193">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4846e-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="4846e-194">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4846e-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="4846e-195">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="4846e-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="4846e-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4846e-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4846e-198">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-198">Type</span></span>

*   <span data-ttu-id="4846e-199">String</span><span class="sxs-lookup"><span data-stu-id="4846e-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4846e-200">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-200">Requirements</span></span>

|<span data-ttu-id="4846e-201">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-201">Requirement</span></span>| <span data-ttu-id="4846e-202">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-203">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-204">1,5</span><span class="sxs-lookup"><span data-stu-id="4846e-204">1.5</span></span> |
|[<span data-ttu-id="4846e-205">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-206">ReadItem</span></span>|
|[<span data-ttu-id="4846e-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4846e-209">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4846e-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4846e-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4846e-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4846e-211">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4846e-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4846e-212">Actuellement, les types d’événement pris `Office.EventType.ItemChanged` en `Office.EventType.OfficeThemeChanged`charge sont et.</span><span class="sxs-lookup"><span data-stu-id="4846e-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-213">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-213">Parameters</span></span>

| <span data-ttu-id="4846e-214">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-214">Name</span></span> | <span data-ttu-id="4846e-215">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-215">Type</span></span> | <span data-ttu-id="4846e-216">Attributs</span><span class="sxs-lookup"><span data-stu-id="4846e-216">Attributes</span></span> | <span data-ttu-id="4846e-217">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4846e-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4846e-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4846e-219">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4846e-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4846e-220">Fonction</span><span class="sxs-lookup"><span data-stu-id="4846e-220">Function</span></span> || <span data-ttu-id="4846e-p105">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4846e-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4846e-224">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-224">Object</span></span> | <span data-ttu-id="4846e-225">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-225">&lt;optional&gt;</span></span> | <span data-ttu-id="4846e-226">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4846e-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4846e-227">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-227">Object</span></span> | <span data-ttu-id="4846e-228">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-228">&lt;optional&gt;</span></span> | <span data-ttu-id="4846e-229">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4846e-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4846e-230">fonction</span><span class="sxs-lookup"><span data-stu-id="4846e-230">function</span></span>| <span data-ttu-id="4846e-231">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-231">&lt;optional&gt;</span></span>|<span data-ttu-id="4846e-232">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4846e-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-233">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-233">Requirements</span></span>

|<span data-ttu-id="4846e-234">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-234">Requirement</span></span>| <span data-ttu-id="4846e-235">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-236">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-237">1,5</span><span class="sxs-lookup"><span data-stu-id="4846e-237">1.5</span></span> |
|[<span data-ttu-id="4846e-238">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-239">ReadItem</span></span> |
|[<span data-ttu-id="4846e-240">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-241">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-242">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-242">Example</span></span>

```javascript
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

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="4846e-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4846e-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4846e-244">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="4846e-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-245">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="4846e-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4846e-p106">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="4846e-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-248">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-248">Parameters</span></span>

|<span data-ttu-id="4846e-249">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-249">Name</span></span>| <span data-ttu-id="4846e-250">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-250">Type</span></span>| <span data-ttu-id="4846e-251">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4846e-252">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-252">String</span></span>|<span data-ttu-id="4846e-253">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="4846e-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="4846e-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4846e-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="4846e-255">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="4846e-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-256">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-256">Requirements</span></span>

|<span data-ttu-id="4846e-257">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-257">Requirement</span></span>| <span data-ttu-id="4846e-258">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-259">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-260">1.3</span><span class="sxs-lookup"><span data-stu-id="4846e-260">1.3</span></span>|
|[<span data-ttu-id="4846e-261">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-262">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4846e-262">Restricted</span></span>|
|[<span data-ttu-id="4846e-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4846e-265">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4846e-265">Returns:</span></span>

<span data-ttu-id="4846e-266">Type : String</span><span class="sxs-lookup"><span data-stu-id="4846e-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4846e-267">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-267">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="4846e-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="4846e-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="4846e-269">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="4846e-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4846e-270">Une application de messagerie pour Outlook sur un ordinateur de bureau ou sur le Web peut utiliser différents fuseaux horaires pour les dates et les heures.</span><span class="sxs-lookup"><span data-stu-id="4846e-270">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="4846e-271">Outlook sur un ordinateur de bureau utilise le fuseau horaire de l’ordinateur client; Outlook sur le Web utilise le fuseau horaire défini dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="4846e-271">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="4846e-272">Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4846e-272">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4846e-273">Si l’application de messagerie est en cours d’exécution dans Outlook sur un `convertToLocalClientTime` client de bureau, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="4846e-273">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="4846e-274">Si l’application de messagerie est en cours d’exécution dans Outlook sur `convertToLocalClientTime` le Web, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire spécifié dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="4846e-274">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-275">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-275">Parameters</span></span>

|<span data-ttu-id="4846e-276">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-276">Name</span></span>| <span data-ttu-id="4846e-277">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-277">Type</span></span>| <span data-ttu-id="4846e-278">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4846e-279">Date</span><span class="sxs-lookup"><span data-stu-id="4846e-279">Date</span></span>|<span data-ttu-id="4846e-280">Objet Date</span><span class="sxs-lookup"><span data-stu-id="4846e-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-281">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-281">Requirements</span></span>

|<span data-ttu-id="4846e-282">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-282">Requirement</span></span>| <span data-ttu-id="4846e-283">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-284">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-285">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-285">1.0</span></span>|
|[<span data-ttu-id="4846e-286">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-287">ReadItem</span></span>|
|[<span data-ttu-id="4846e-288">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-289">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4846e-290">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4846e-290">Returns:</span></span>

<span data-ttu-id="4846e-291">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="4846e-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="4846e-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4846e-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4846e-293">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="4846e-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-294">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="4846e-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4846e-p109">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="4846e-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-297">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-297">Parameters</span></span>

|<span data-ttu-id="4846e-298">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-298">Name</span></span>| <span data-ttu-id="4846e-299">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-299">Type</span></span>| <span data-ttu-id="4846e-300">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4846e-301">String</span><span class="sxs-lookup"><span data-stu-id="4846e-301">String</span></span>|<span data-ttu-id="4846e-302">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="4846e-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="4846e-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4846e-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="4846e-304">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="4846e-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-305">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-305">Requirements</span></span>

|<span data-ttu-id="4846e-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-306">Requirement</span></span>| <span data-ttu-id="4846e-307">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-308">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-309">1.3</span><span class="sxs-lookup"><span data-stu-id="4846e-309">1.3</span></span>|
|[<span data-ttu-id="4846e-310">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-311">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4846e-311">Restricted</span></span>|
|[<span data-ttu-id="4846e-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-313">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4846e-314">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4846e-314">Returns:</span></span>

<span data-ttu-id="4846e-315">Type : String</span><span class="sxs-lookup"><span data-stu-id="4846e-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4846e-316">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-316">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4846e-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4846e-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4846e-318">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="4846e-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4846e-319">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="4846e-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-320">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-320">Parameters</span></span>

|<span data-ttu-id="4846e-321">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-321">Name</span></span>| <span data-ttu-id="4846e-322">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-322">Type</span></span>| <span data-ttu-id="4846e-323">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4846e-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4846e-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="4846e-325">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="4846e-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-326">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-326">Requirements</span></span>

|<span data-ttu-id="4846e-327">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-327">Requirement</span></span>| <span data-ttu-id="4846e-328">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-329">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-330">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-330">1.0</span></span>|
|[<span data-ttu-id="4846e-331">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-332">ReadItem</span></span>|
|[<span data-ttu-id="4846e-333">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-334">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4846e-335">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4846e-335">Returns:</span></span>

<span data-ttu-id="4846e-336">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="4846e-336">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="4846e-337">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4846e-337">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4846e-338">Date</span><span class="sxs-lookup"><span data-stu-id="4846e-338">Date</span></span></dd>

</dl>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="4846e-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4846e-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4846e-340">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4846e-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-341">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="4846e-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4846e-342">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4846e-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4846e-343">Dans Outlook sur Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série.</span><span class="sxs-lookup"><span data-stu-id="4846e-343">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="4846e-344">En effet, dans Outlook sur Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="4846e-344">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4846e-345">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32KO nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="4846e-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4846e-346">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4846e-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-347">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-347">Parameters</span></span>

|<span data-ttu-id="4846e-348">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-348">Name</span></span>| <span data-ttu-id="4846e-349">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-349">Type</span></span>| <span data-ttu-id="4846e-350">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4846e-351">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-351">String</span></span>|<span data-ttu-id="4846e-352">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4846e-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-353">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-353">Requirements</span></span>

|<span data-ttu-id="4846e-354">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-354">Requirement</span></span>| <span data-ttu-id="4846e-355">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-356">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-357">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-357">1.0</span></span>|
|[<span data-ttu-id="4846e-358">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-359">ReadItem</span></span>|
|[<span data-ttu-id="4846e-360">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-361">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-362">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-362">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="4846e-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4846e-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4846e-364">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="4846e-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-365">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="4846e-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4846e-366">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4846e-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4846e-367">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32 Ko nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="4846e-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4846e-368">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4846e-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4846e-p111">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4846e-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-371">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-371">Parameters</span></span>

|<span data-ttu-id="4846e-372">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-372">Name</span></span>| <span data-ttu-id="4846e-373">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-373">Type</span></span>| <span data-ttu-id="4846e-374">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4846e-375">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-375">String</span></span>|<span data-ttu-id="4846e-376">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="4846e-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-377">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-377">Requirements</span></span>

|<span data-ttu-id="4846e-378">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-378">Requirement</span></span>| <span data-ttu-id="4846e-379">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-380">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-381">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-381">1.0</span></span>|
|[<span data-ttu-id="4846e-382">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-383">ReadItem</span></span>|
|[<span data-ttu-id="4846e-384">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-385">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-386">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-386">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4846e-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4846e-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4846e-388">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="4846e-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-389">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="4846e-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4846e-p112">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="4846e-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4846e-392">Dans Outlook sur le Web et les appareils mobiles, cette méthode affiche toujours un formulaire avec un champ participants.</span><span class="sxs-lookup"><span data-stu-id="4846e-392">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="4846e-393">Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="4846e-393">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="4846e-394">Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="4846e-394">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4846e-p114">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="4846e-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4846e-397">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="4846e-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-398">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-399">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="4846e-399">All parameters are optional.</span></span>

|<span data-ttu-id="4846e-400">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-400">Name</span></span>| <span data-ttu-id="4846e-401">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-401">Type</span></span>| <span data-ttu-id="4846e-402">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4846e-403">Object</span><span class="sxs-lookup"><span data-stu-id="4846e-403">Object</span></span> | <span data-ttu-id="4846e-404">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4846e-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4846e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4846e-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4846e-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4846e-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4846e-411">Date</span><span class="sxs-lookup"><span data-stu-id="4846e-411">Date</span></span> | <span data-ttu-id="4846e-412">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4846e-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4846e-413">Date</span><span class="sxs-lookup"><span data-stu-id="4846e-413">Date</span></span> | <span data-ttu-id="4846e-414">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4846e-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4846e-415">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-415">String</span></span> | <span data-ttu-id="4846e-p117">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4846e-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4846e-p118">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4846e-421">String</span><span class="sxs-lookup"><span data-stu-id="4846e-421">String</span></span> | <span data-ttu-id="4846e-p119">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4846e-424">String</span><span class="sxs-lookup"><span data-stu-id="4846e-424">String</span></span> | <span data-ttu-id="4846e-p120">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4846e-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4846e-427">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-427">Requirements</span></span>

|<span data-ttu-id="4846e-428">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-428">Requirement</span></span>| <span data-ttu-id="4846e-429">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-430">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-431">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-431">1.0</span></span>|
|[<span data-ttu-id="4846e-432">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-433">ReadItem</span></span>|
|[<span data-ttu-id="4846e-434">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-435">Lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-436">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-436">Example</span></span>

```javascript
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

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="4846e-437">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="4846e-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="4846e-438">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="4846e-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="4846e-439">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="4846e-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="4846e-440">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="4846e-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4846e-441">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="4846e-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-442">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-443">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="4846e-443">All parameters are optional.</span></span>

|<span data-ttu-id="4846e-444">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-444">Name</span></span>| <span data-ttu-id="4846e-445">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-445">Type</span></span>| <span data-ttu-id="4846e-446">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4846e-447">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-447">Object</span></span> | <span data-ttu-id="4846e-448">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="4846e-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="4846e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4846e-450">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="4846e-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="4846e-451">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="4846e-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4846e-453">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="4846e-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="4846e-454">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="4846e-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4846e-456">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="4846e-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="4846e-457">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4846e-458">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-458">String</span></span> | <span data-ttu-id="4846e-459">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="4846e-459">A string containing the subject of the message.</span></span> <span data-ttu-id="4846e-460">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4846e-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="4846e-461">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-461">String</span></span> | <span data-ttu-id="4846e-462">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="4846e-462">The HTML body of the message.</span></span> <span data-ttu-id="4846e-463">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4846e-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="4846e-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4846e-465">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4846e-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="4846e-466">String</span><span class="sxs-lookup"><span data-stu-id="4846e-466">String</span></span> | <span data-ttu-id="4846e-p127">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4846e-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="4846e-469">String</span><span class="sxs-lookup"><span data-stu-id="4846e-469">String</span></span> | <span data-ttu-id="4846e-470">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4846e-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="4846e-471">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4846e-471">String</span></span> | <span data-ttu-id="4846e-p128">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4846e-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="4846e-474">Booléen</span><span class="sxs-lookup"><span data-stu-id="4846e-474">Boolean</span></span> | <span data-ttu-id="4846e-p129">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4846e-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="4846e-477">String</span><span class="sxs-lookup"><span data-stu-id="4846e-477">String</span></span> | <span data-ttu-id="4846e-478">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="4846e-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="4846e-479">ID d’élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="4846e-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="4846e-480">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4846e-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="4846e-481">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-481">Requirements</span></span>

|<span data-ttu-id="4846e-482">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-482">Requirement</span></span>| <span data-ttu-id="4846e-483">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-484">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-485">1.6</span><span class="sxs-lookup"><span data-stu-id="4846e-485">1.6</span></span> |
|[<span data-ttu-id="4846e-486">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-487">ReadItem</span></span>|
|[<span data-ttu-id="4846e-488">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-489">Lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-490">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-490">Example</span></span>

```javascript
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

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="4846e-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4846e-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="4846e-492">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4846e-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="4846e-p131">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="4846e-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-495">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4846e-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="4846e-496">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="4846e-496">**REST Tokens**</span></span>

<span data-ttu-id="4846e-p132">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="4846e-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="4846e-500">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="4846e-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="4846e-501">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="4846e-501">**EWS Tokens**</span></span>

<span data-ttu-id="4846e-p133">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4846e-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="4846e-504">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="4846e-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-505">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-505">Parameters</span></span>

|<span data-ttu-id="4846e-506">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-506">Name</span></span>| <span data-ttu-id="4846e-507">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-507">Type</span></span>| <span data-ttu-id="4846e-508">Attributs</span><span class="sxs-lookup"><span data-stu-id="4846e-508">Attributes</span></span>| <span data-ttu-id="4846e-509">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-509">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="4846e-510">Object</span><span class="sxs-lookup"><span data-stu-id="4846e-510">Object</span></span> | <span data-ttu-id="4846e-511">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-511">&lt;optional&gt;</span></span> | <span data-ttu-id="4846e-512">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4846e-512">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="4846e-513">Boolean</span><span class="sxs-lookup"><span data-stu-id="4846e-513">Boolean</span></span> |  <span data-ttu-id="4846e-514">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-514">&lt;optional&gt;</span></span> | <span data-ttu-id="4846e-p134">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="4846e-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4846e-517">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-517">Object</span></span> |  <span data-ttu-id="4846e-518">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-518">&lt;optional&gt;</span></span> | <span data-ttu-id="4846e-519">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4846e-519">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="4846e-520">fonction</span><span class="sxs-lookup"><span data-stu-id="4846e-520">function</span></span>||<span data-ttu-id="4846e-521">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4846e-521">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4846e-522">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4846e-522">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4846e-523">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4846e-523">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4846e-524">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4846e-524">Errors</span></span>

|<span data-ttu-id="4846e-525">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4846e-525">Error code</span></span>|<span data-ttu-id="4846e-526">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-526">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4846e-527">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4846e-527">The request has failed.</span></span> <span data-ttu-id="4846e-528">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="4846e-528">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4846e-529">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4846e-529">The Exchange server returned an error.</span></span> <span data-ttu-id="4846e-530">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4846e-530">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4846e-531">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4846e-531">The user is no longer connected to the network.</span></span> <span data-ttu-id="4846e-532">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="4846e-532">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-533">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-533">Requirements</span></span>

|<span data-ttu-id="4846e-534">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-534">Requirement</span></span>| <span data-ttu-id="4846e-535">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-535">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-536">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-536">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-537">1,5</span><span class="sxs-lookup"><span data-stu-id="4846e-537">1.5</span></span> |
|[<span data-ttu-id="4846e-538">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-538">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-539">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-539">ReadItem</span></span>|
|[<span data-ttu-id="4846e-540">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-540">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-541">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-541">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-542">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-542">Example</span></span>

```javascript
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

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4846e-543">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4846e-543">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4846e-544">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="4846e-544">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4846e-p138">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="4846e-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4846e-p139">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4846e-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4846e-550">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="4846e-550">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="4846e-p140">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4846e-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-553">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-553">Parameters</span></span>

|<span data-ttu-id="4846e-554">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-554">Name</span></span>| <span data-ttu-id="4846e-555">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-555">Type</span></span>| <span data-ttu-id="4846e-556">Attributs</span><span class="sxs-lookup"><span data-stu-id="4846e-556">Attributes</span></span>| <span data-ttu-id="4846e-557">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-557">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4846e-558">function</span><span class="sxs-lookup"><span data-stu-id="4846e-558">function</span></span>||<span data-ttu-id="4846e-559">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4846e-559">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4846e-560">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4846e-560">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4846e-561">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4846e-561">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4846e-562">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-562">Object</span></span>| <span data-ttu-id="4846e-563">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-563">&lt;optional&gt;</span></span>|<span data-ttu-id="4846e-564">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4846e-564">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4846e-565">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4846e-565">Errors</span></span>

|<span data-ttu-id="4846e-566">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4846e-566">Error code</span></span>|<span data-ttu-id="4846e-567">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-567">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4846e-568">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4846e-568">The request has failed.</span></span> <span data-ttu-id="4846e-569">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="4846e-569">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4846e-570">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4846e-570">The Exchange server returned an error.</span></span> <span data-ttu-id="4846e-571">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4846e-571">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4846e-572">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4846e-572">The user is no longer connected to the network.</span></span> <span data-ttu-id="4846e-573">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="4846e-573">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-574">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-574">Requirements</span></span>

|<span data-ttu-id="4846e-575">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-575">Requirement</span></span>| <span data-ttu-id="4846e-576">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-576">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-577">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-577">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-578">1.3</span><span class="sxs-lookup"><span data-stu-id="4846e-578">1.3</span></span>|
|[<span data-ttu-id="4846e-579">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-579">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-580">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-580">ReadItem</span></span>|
|[<span data-ttu-id="4846e-581">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-581">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-582">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-582">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-583">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-583">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4846e-584">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4846e-584">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4846e-585">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="4846e-585">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4846e-586">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="4846e-586">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-587">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-587">Parameters</span></span>

|<span data-ttu-id="4846e-588">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-588">Name</span></span>| <span data-ttu-id="4846e-589">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-589">Type</span></span>| <span data-ttu-id="4846e-590">Attributs</span><span class="sxs-lookup"><span data-stu-id="4846e-590">Attributes</span></span>| <span data-ttu-id="4846e-591">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-591">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4846e-592">fonction</span><span class="sxs-lookup"><span data-stu-id="4846e-592">function</span></span>||<span data-ttu-id="4846e-593">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4846e-593">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4846e-594">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4846e-594">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4846e-595">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4846e-595">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4846e-596">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-596">Object</span></span>| <span data-ttu-id="4846e-597">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-597">&lt;optional&gt;</span></span>|<span data-ttu-id="4846e-598">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4846e-598">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4846e-599">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4846e-599">Errors</span></span>

|<span data-ttu-id="4846e-600">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4846e-600">Error code</span></span>|<span data-ttu-id="4846e-601">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-601">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4846e-602">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4846e-602">The request has failed.</span></span> <span data-ttu-id="4846e-603">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="4846e-603">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4846e-604">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4846e-604">The Exchange server returned an error.</span></span> <span data-ttu-id="4846e-605">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4846e-605">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4846e-606">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4846e-606">The user is no longer connected to the network.</span></span> <span data-ttu-id="4846e-607">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="4846e-607">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-608">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-608">Requirements</span></span>

|<span data-ttu-id="4846e-609">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-609">Requirement</span></span>| <span data-ttu-id="4846e-610">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-610">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-611">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-611">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-612">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-612">1.0</span></span>|
|[<span data-ttu-id="4846e-613">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-613">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-614">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-614">ReadItem</span></span>|
|[<span data-ttu-id="4846e-615">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-615">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-616">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-616">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-617">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-617">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4846e-618">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4846e-618">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4846e-619">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4846e-619">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-620">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="4846e-620">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4846e-621">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="4846e-621">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="4846e-622">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="4846e-622">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4846e-623">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4846e-623">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4846e-624">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="4846e-624">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4846e-625">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="4846e-625">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4846e-626">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="4846e-626">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4846e-627">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="4846e-627">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4846e-p148">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="4846e-p148">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4846e-630">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="4846e-630">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4846e-631">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="4846e-631">Version differences</span></span>

<span data-ttu-id="4846e-632">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="4846e-632">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4846e-633">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage.</span><span class="sxs-lookup"><span data-stu-id="4846e-633">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="4846e-634">Vous pouvez déterminer si votre application de messagerie est en cours d’exécution dans Outlook sur le Web ou sur un client de bureau à l’aide de la propriété Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="4846e-634">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="4846e-635">Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="4846e-635">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-636">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-636">Parameters</span></span>

|<span data-ttu-id="4846e-637">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-637">Name</span></span>| <span data-ttu-id="4846e-638">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-638">Type</span></span>| <span data-ttu-id="4846e-639">Attributs</span><span class="sxs-lookup"><span data-stu-id="4846e-639">Attributes</span></span>| <span data-ttu-id="4846e-640">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-640">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4846e-641">String</span><span class="sxs-lookup"><span data-stu-id="4846e-641">String</span></span>||<span data-ttu-id="4846e-642">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="4846e-642">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4846e-643">function</span><span class="sxs-lookup"><span data-stu-id="4846e-643">function</span></span>||<span data-ttu-id="4846e-644">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4846e-644">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4846e-645">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4846e-645">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4846e-646">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="4846e-646">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="4846e-647">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-647">Object</span></span>| <span data-ttu-id="4846e-648">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-648">&lt;optional&gt;</span></span>|<span data-ttu-id="4846e-649">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4846e-649">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-650">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-650">Requirements</span></span>

|<span data-ttu-id="4846e-651">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-651">Requirement</span></span>| <span data-ttu-id="4846e-652">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-652">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-653">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-653">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-654">1.0</span><span class="sxs-lookup"><span data-stu-id="4846e-654">1.0</span></span>|
|[<span data-ttu-id="4846e-655">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-655">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-656">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4846e-656">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4846e-657">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-657">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-658">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-658">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4846e-659">Exemple</span><span class="sxs-lookup"><span data-stu-id="4846e-659">Example</span></span>

<span data-ttu-id="4846e-660">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4846e-660">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
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

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4846e-661">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4846e-661">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4846e-662">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4846e-662">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4846e-663">Actuellement, les types d’événement pris `Office.EventType.ItemChanged` en `Office.EventType.OfficeThemeChanged`charge sont et.</span><span class="sxs-lookup"><span data-stu-id="4846e-663">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4846e-664">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4846e-664">Parameters</span></span>

| <span data-ttu-id="4846e-665">Nom</span><span class="sxs-lookup"><span data-stu-id="4846e-665">Name</span></span> | <span data-ttu-id="4846e-666">Type</span><span class="sxs-lookup"><span data-stu-id="4846e-666">Type</span></span> | <span data-ttu-id="4846e-667">Attributs</span><span class="sxs-lookup"><span data-stu-id="4846e-667">Attributes</span></span> | <span data-ttu-id="4846e-668">Description</span><span class="sxs-lookup"><span data-stu-id="4846e-668">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4846e-669">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4846e-669">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4846e-670">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4846e-670">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="4846e-671">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-671">Object</span></span> | <span data-ttu-id="4846e-672">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-672">&lt;optional&gt;</span></span> | <span data-ttu-id="4846e-673">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4846e-673">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4846e-674">Objet</span><span class="sxs-lookup"><span data-stu-id="4846e-674">Object</span></span> | <span data-ttu-id="4846e-675">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-675">&lt;optional&gt;</span></span> | <span data-ttu-id="4846e-676">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4846e-676">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4846e-677">fonction</span><span class="sxs-lookup"><span data-stu-id="4846e-677">function</span></span>| <span data-ttu-id="4846e-678">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4846e-678">&lt;optional&gt;</span></span>|<span data-ttu-id="4846e-679">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4846e-679">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4846e-680">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4846e-680">Requirements</span></span>

|<span data-ttu-id="4846e-681">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4846e-681">Requirement</span></span>| <span data-ttu-id="4846e-682">Valeur</span><span class="sxs-lookup"><span data-stu-id="4846e-682">Value</span></span>|
|---|---|
|[<span data-ttu-id="4846e-683">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4846e-683">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4846e-684">1,5</span><span class="sxs-lookup"><span data-stu-id="4846e-684">1.5</span></span> |
|[<span data-ttu-id="4846e-685">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4846e-685">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4846e-686">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4846e-686">ReadItem</span></span> |
|[<span data-ttu-id="4846e-687">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4846e-687">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4846e-688">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4846e-688">Compose or Read</span></span>|
