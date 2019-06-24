---
title: Office. Context. Mailbox-ensemble de conditions requises 1,7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 95fb4ce6bcc3c44c77dc4623a12b140ca979949c
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127148"
---
# <a name="mailbox"></a><span data-ttu-id="236b7-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="236b7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="236b7-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="236b7-104">Permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="236b7-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="236b7-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-105">Requirements</span></span>

|<span data-ttu-id="236b7-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-106">Requirement</span></span>| <span data-ttu-id="236b7-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-109">1.0</span></span>|
|[<span data-ttu-id="236b7-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="236b7-111">Restricted</span></span>|
|[<span data-ttu-id="236b7-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="236b7-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="236b7-114">Members and methods</span></span>

| <span data-ttu-id="236b7-115">Membre</span><span class="sxs-lookup"><span data-stu-id="236b7-115">Member</span></span> | <span data-ttu-id="236b7-116">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="236b7-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="236b7-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="236b7-118">Membre</span><span class="sxs-lookup"><span data-stu-id="236b7-118">Member</span></span> |
| [<span data-ttu-id="236b7-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="236b7-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="236b7-120">Membre</span><span class="sxs-lookup"><span data-stu-id="236b7-120">Member</span></span> |
| [<span data-ttu-id="236b7-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="236b7-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="236b7-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-122">Method</span></span> |
| [<span data-ttu-id="236b7-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="236b7-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="236b7-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-124">Method</span></span> |
| [<span data-ttu-id="236b7-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="236b7-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="236b7-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-126">Method</span></span> |
| [<span data-ttu-id="236b7-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="236b7-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="236b7-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-128">Method</span></span> |
| [<span data-ttu-id="236b7-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="236b7-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="236b7-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-130">Method</span></span> |
| [<span data-ttu-id="236b7-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="236b7-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="236b7-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-132">Method</span></span> |
| [<span data-ttu-id="236b7-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="236b7-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="236b7-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-134">Method</span></span> |
| [<span data-ttu-id="236b7-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="236b7-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="236b7-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-136">Method</span></span> |
| [<span data-ttu-id="236b7-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="236b7-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="236b7-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-138">Method</span></span> |
| [<span data-ttu-id="236b7-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="236b7-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="236b7-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-140">Method</span></span> |
| [<span data-ttu-id="236b7-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="236b7-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="236b7-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-142">Method</span></span> |
| [<span data-ttu-id="236b7-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="236b7-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="236b7-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-144">Method</span></span> |
| [<span data-ttu-id="236b7-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="236b7-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="236b7-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-146">Method</span></span> |
| [<span data-ttu-id="236b7-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="236b7-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="236b7-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="236b7-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="236b7-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="236b7-149">Namespaces</span></span>

<span data-ttu-id="236b7-150">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="236b7-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="236b7-151">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="236b7-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="236b7-152">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="236b7-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="236b7-153">Membres</span><span class="sxs-lookup"><span data-stu-id="236b7-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="236b7-154">ewsUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-154">ewsUrl: String</span></span>

<span data-ttu-id="236b7-155">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="236b7-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="236b7-156">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="236b7-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-157">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="236b7-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="236b7-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="236b7-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="236b7-160">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="236b7-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="236b7-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="236b7-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="236b7-163">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-163">Type</span></span>

*   <span data-ttu-id="236b7-164">String</span><span class="sxs-lookup"><span data-stu-id="236b7-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="236b7-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-165">Requirements</span></span>

|<span data-ttu-id="236b7-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-166">Requirement</span></span>| <span data-ttu-id="236b7-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-169">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-169">1.0</span></span>|
|[<span data-ttu-id="236b7-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-171">ReadItem</span></span>|
|[<span data-ttu-id="236b7-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="236b7-174">restUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-174">restUrl: String</span></span>

<span data-ttu-id="236b7-175">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="236b7-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="236b7-176">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="236b7-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="236b7-177">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="236b7-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="236b7-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="236b7-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="236b7-180">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-180">Type</span></span>

*   <span data-ttu-id="236b7-181">String</span><span class="sxs-lookup"><span data-stu-id="236b7-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="236b7-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-182">Requirements</span></span>

|<span data-ttu-id="236b7-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-183">Requirement</span></span>| <span data-ttu-id="236b7-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-186">1,5</span><span class="sxs-lookup"><span data-stu-id="236b7-186">1.5</span></span> |
|[<span data-ttu-id="236b7-187">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-188">ReadItem</span></span>|
|[<span data-ttu-id="236b7-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="236b7-191">Méthodes</span><span class="sxs-lookup"><span data-stu-id="236b7-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="236b7-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="236b7-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="236b7-193">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="236b7-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="236b7-194">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="236b7-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-195">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-195">Parameters</span></span>

| <span data-ttu-id="236b7-196">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-196">Name</span></span> | <span data-ttu-id="236b7-197">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-197">Type</span></span> | <span data-ttu-id="236b7-198">Attributs</span><span class="sxs-lookup"><span data-stu-id="236b7-198">Attributes</span></span> | <span data-ttu-id="236b7-199">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="236b7-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="236b7-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="236b7-201">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="236b7-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="236b7-202">Fonction</span><span class="sxs-lookup"><span data-stu-id="236b7-202">Function</span></span> || <span data-ttu-id="236b7-p105">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="236b7-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="236b7-206">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-206">Object</span></span> | <span data-ttu-id="236b7-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-207">&lt;optional&gt;</span></span> | <span data-ttu-id="236b7-208">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="236b7-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="236b7-209">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-209">Object</span></span> | <span data-ttu-id="236b7-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-210">&lt;optional&gt;</span></span> | <span data-ttu-id="236b7-211">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="236b7-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="236b7-212">fonction</span><span class="sxs-lookup"><span data-stu-id="236b7-212">function</span></span>| <span data-ttu-id="236b7-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-213">&lt;optional&gt;</span></span>|<span data-ttu-id="236b7-214">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="236b7-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-215">Requirements</span></span>

|<span data-ttu-id="236b7-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-216">Requirement</span></span>| <span data-ttu-id="236b7-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-219">1,5</span><span class="sxs-lookup"><span data-stu-id="236b7-219">1.5</span></span> |
|[<span data-ttu-id="236b7-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-221">ReadItem</span></span> |
|[<span data-ttu-id="236b7-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-223">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-224">Example</span></span>

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
};
```

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="236b7-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="236b7-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="236b7-226">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="236b7-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-227">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="236b7-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="236b7-p106">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="236b7-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-230">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-230">Parameters</span></span>

|<span data-ttu-id="236b7-231">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-231">Name</span></span>| <span data-ttu-id="236b7-232">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-232">Type</span></span>| <span data-ttu-id="236b7-233">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="236b7-234">String</span><span class="sxs-lookup"><span data-stu-id="236b7-234">String</span></span>|<span data-ttu-id="236b7-235">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="236b7-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="236b7-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="236b7-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="236b7-237">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="236b7-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-238">Requirements</span></span>

|<span data-ttu-id="236b7-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-239">Requirement</span></span>| <span data-ttu-id="236b7-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-242">1.3</span><span class="sxs-lookup"><span data-stu-id="236b7-242">1.3</span></span>|
|[<span data-ttu-id="236b7-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-244">Restreinte</span><span class="sxs-lookup"><span data-stu-id="236b7-244">Restricted</span></span>|
|[<span data-ttu-id="236b7-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="236b7-247">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="236b7-247">Returns:</span></span>

<span data-ttu-id="236b7-248">Type : String</span><span class="sxs-lookup"><span data-stu-id="236b7-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="236b7-249">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="236b7-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="236b7-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="236b7-251">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="236b7-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="236b7-252">Une application de messagerie pour Outlook sur un ordinateur de bureau ou sur le Web peut utiliser différents fuseaux horaires pour les dates et les heures.</span><span class="sxs-lookup"><span data-stu-id="236b7-252">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="236b7-253">Outlook sur un ordinateur de bureau utilise le fuseau horaire de l’ordinateur client; Outlook sur le Web utilise le fuseau horaire défini dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="236b7-253">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="236b7-254">Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="236b7-254">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="236b7-255">Si l’application de messagerie est en cours d’exécution dans Outlook sur un `convertToLocalClientTime` client de bureau, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="236b7-255">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="236b7-256">Si l’application de messagerie est en cours d’exécution dans Outlook sur `convertToLocalClientTime` le Web, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire spécifié dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="236b7-256">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-257">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-257">Parameters</span></span>

|<span data-ttu-id="236b7-258">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-258">Name</span></span>| <span data-ttu-id="236b7-259">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-259">Type</span></span>| <span data-ttu-id="236b7-260">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="236b7-261">Date</span><span class="sxs-lookup"><span data-stu-id="236b7-261">Date</span></span>|<span data-ttu-id="236b7-262">Objet Date</span><span class="sxs-lookup"><span data-stu-id="236b7-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-263">Requirements</span></span>

|<span data-ttu-id="236b7-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-264">Requirement</span></span>| <span data-ttu-id="236b7-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-267">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-267">1.0</span></span>|
|[<span data-ttu-id="236b7-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-269">ReadItem</span></span>|
|[<span data-ttu-id="236b7-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-271">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="236b7-272">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="236b7-272">Returns:</span></span>

<span data-ttu-id="236b7-273">Type : [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="236b7-273">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="236b7-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="236b7-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="236b7-275">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="236b7-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-276">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="236b7-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="236b7-p109">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="236b7-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-279">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-279">Parameters</span></span>

|<span data-ttu-id="236b7-280">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-280">Name</span></span>| <span data-ttu-id="236b7-281">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-281">Type</span></span>| <span data-ttu-id="236b7-282">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="236b7-283">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-283">String</span></span>|<span data-ttu-id="236b7-284">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="236b7-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="236b7-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="236b7-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="236b7-286">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="236b7-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-287">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-287">Requirements</span></span>

|<span data-ttu-id="236b7-288">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-288">Requirement</span></span>| <span data-ttu-id="236b7-289">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-290">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-291">1.3</span><span class="sxs-lookup"><span data-stu-id="236b7-291">1.3</span></span>|
|[<span data-ttu-id="236b7-292">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-293">Restreinte</span><span class="sxs-lookup"><span data-stu-id="236b7-293">Restricted</span></span>|
|[<span data-ttu-id="236b7-294">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-295">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="236b7-296">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="236b7-296">Returns:</span></span>

<span data-ttu-id="236b7-297">Type : String</span><span class="sxs-lookup"><span data-stu-id="236b7-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="236b7-298">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="236b7-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="236b7-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="236b7-300">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="236b7-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="236b7-301">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="236b7-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-302">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-302">Parameters</span></span>

|<span data-ttu-id="236b7-303">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-303">Name</span></span>| <span data-ttu-id="236b7-304">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-304">Type</span></span>| <span data-ttu-id="236b7-305">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="236b7-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="236b7-306">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="236b7-307">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="236b7-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-308">Requirements</span></span>

|<span data-ttu-id="236b7-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-309">Requirement</span></span>| <span data-ttu-id="236b7-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-312">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-312">1.0</span></span>|
|[<span data-ttu-id="236b7-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-314">ReadItem</span></span>|
|[<span data-ttu-id="236b7-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-316">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="236b7-317">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="236b7-317">Returns:</span></span>

<span data-ttu-id="236b7-318">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="236b7-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="236b7-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="236b7-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="236b7-320">Date</span><span class="sxs-lookup"><span data-stu-id="236b7-320">Date</span></span></dd>

</dl>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="236b7-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="236b7-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="236b7-322">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="236b7-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-323">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="236b7-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="236b7-324">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="236b7-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="236b7-325">Dans Outlook sur Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série.</span><span class="sxs-lookup"><span data-stu-id="236b7-325">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="236b7-326">En effet, dans Outlook sur Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="236b7-326">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="236b7-327">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32KO nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="236b7-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="236b7-328">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="236b7-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-329">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-329">Parameters</span></span>

|<span data-ttu-id="236b7-330">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-330">Name</span></span>| <span data-ttu-id="236b7-331">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-331">Type</span></span>| <span data-ttu-id="236b7-332">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="236b7-333">String</span><span class="sxs-lookup"><span data-stu-id="236b7-333">String</span></span>|<span data-ttu-id="236b7-334">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="236b7-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-335">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-335">Requirements</span></span>

|<span data-ttu-id="236b7-336">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-336">Requirement</span></span>| <span data-ttu-id="236b7-337">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-338">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-339">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-339">1.0</span></span>|
|[<span data-ttu-id="236b7-340">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-341">ReadItem</span></span>|
|[<span data-ttu-id="236b7-342">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-343">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-344">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="236b7-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="236b7-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="236b7-346">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="236b7-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-347">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="236b7-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="236b7-348">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="236b7-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="236b7-349">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32 Ko nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="236b7-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="236b7-350">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="236b7-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="236b7-p111">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="236b7-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-353">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-353">Parameters</span></span>

|<span data-ttu-id="236b7-354">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-354">Name</span></span>| <span data-ttu-id="236b7-355">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-355">Type</span></span>| <span data-ttu-id="236b7-356">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="236b7-357">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-357">String</span></span>|<span data-ttu-id="236b7-358">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="236b7-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-359">Requirements</span></span>

|<span data-ttu-id="236b7-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-360">Requirement</span></span>| <span data-ttu-id="236b7-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-363">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-363">1.0</span></span>|
|[<span data-ttu-id="236b7-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-365">ReadItem</span></span>|
|[<span data-ttu-id="236b7-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-367">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-368">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="236b7-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="236b7-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="236b7-370">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="236b7-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-371">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="236b7-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="236b7-p112">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="236b7-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="236b7-374">Dans Outlook sur le Web et les appareils mobiles, cette méthode affiche toujours un formulaire avec un champ participants.</span><span class="sxs-lookup"><span data-stu-id="236b7-374">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="236b7-375">Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="236b7-375">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="236b7-376">Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="236b7-376">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="236b7-p114">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="236b7-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="236b7-379">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="236b7-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-380">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-381">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="236b7-381">All parameters are optional.</span></span>

|<span data-ttu-id="236b7-382">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-382">Name</span></span>| <span data-ttu-id="236b7-383">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-383">Type</span></span>| <span data-ttu-id="236b7-384">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="236b7-385">Object</span><span class="sxs-lookup"><span data-stu-id="236b7-385">Object</span></span> | <span data-ttu-id="236b7-386">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="236b7-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="236b7-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="236b7-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="236b7-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="236b7-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="236b7-393">Date</span><span class="sxs-lookup"><span data-stu-id="236b7-393">Date</span></span> | <span data-ttu-id="236b7-394">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="236b7-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="236b7-395">Date</span><span class="sxs-lookup"><span data-stu-id="236b7-395">Date</span></span> | <span data-ttu-id="236b7-396">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="236b7-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="236b7-397">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-397">String</span></span> | <span data-ttu-id="236b7-p117">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="236b7-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="236b7-p118">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="236b7-403">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-403">String</span></span> | <span data-ttu-id="236b7-p119">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="236b7-406">String</span><span class="sxs-lookup"><span data-stu-id="236b7-406">String</span></span> | <span data-ttu-id="236b7-p120">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="236b7-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="236b7-409">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-409">Requirements</span></span>

|<span data-ttu-id="236b7-410">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-410">Requirement</span></span>| <span data-ttu-id="236b7-411">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-412">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-413">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-413">1.0</span></span>|
|[<span data-ttu-id="236b7-414">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-415">ReadItem</span></span>|
|[<span data-ttu-id="236b7-416">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-417">Lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-418">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="236b7-419">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="236b7-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="236b7-420">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="236b7-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="236b7-421">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="236b7-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="236b7-422">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="236b7-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="236b7-423">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="236b7-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-424">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-425">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="236b7-425">All parameters are optional.</span></span>

|<span data-ttu-id="236b7-426">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-426">Name</span></span>| <span data-ttu-id="236b7-427">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-427">Type</span></span>| <span data-ttu-id="236b7-428">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="236b7-429">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-429">Object</span></span> | <span data-ttu-id="236b7-430">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="236b7-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="236b7-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="236b7-432">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="236b7-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="236b7-433">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="236b7-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="236b7-435">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="236b7-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="236b7-436">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="236b7-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="236b7-438">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="236b7-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="236b7-439">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="236b7-440">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-440">String</span></span> | <span data-ttu-id="236b7-441">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="236b7-441">A string containing the subject of the message.</span></span> <span data-ttu-id="236b7-442">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="236b7-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="236b7-443">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-443">String</span></span> | <span data-ttu-id="236b7-444">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="236b7-444">The HTML body of the message.</span></span> <span data-ttu-id="236b7-445">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="236b7-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="236b7-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="236b7-447">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="236b7-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="236b7-448">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-448">String</span></span> | <span data-ttu-id="236b7-p127">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="236b7-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="236b7-451">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-451">String</span></span> | <span data-ttu-id="236b7-452">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="236b7-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="236b7-453">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-453">String</span></span> | <span data-ttu-id="236b7-p128">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="236b7-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="236b7-456">Booléen</span><span class="sxs-lookup"><span data-stu-id="236b7-456">Boolean</span></span> | <span data-ttu-id="236b7-p129">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="236b7-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="236b7-459">Chaîne</span><span class="sxs-lookup"><span data-stu-id="236b7-459">String</span></span> | <span data-ttu-id="236b7-460">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="236b7-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="236b7-461">ID d’élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="236b7-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="236b7-462">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="236b7-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="236b7-463">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-463">Requirements</span></span>

|<span data-ttu-id="236b7-464">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-464">Requirement</span></span>| <span data-ttu-id="236b7-465">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-466">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-467">1.6</span><span class="sxs-lookup"><span data-stu-id="236b7-467">1.6</span></span> |
|[<span data-ttu-id="236b7-468">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-469">ReadItem</span></span>|
|[<span data-ttu-id="236b7-470">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-471">Lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-472">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-472">Example</span></span>

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="236b7-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="236b7-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="236b7-474">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="236b7-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="236b7-p131">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="236b7-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-477">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="236b7-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="236b7-478">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="236b7-478">**REST Tokens**</span></span>

<span data-ttu-id="236b7-p132">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="236b7-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="236b7-482">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="236b7-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="236b7-483">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="236b7-483">**EWS Tokens**</span></span>

<span data-ttu-id="236b7-p133">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="236b7-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="236b7-486">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="236b7-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-487">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-487">Parameters</span></span>

|<span data-ttu-id="236b7-488">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-488">Name</span></span>| <span data-ttu-id="236b7-489">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-489">Type</span></span>| <span data-ttu-id="236b7-490">Attributs</span><span class="sxs-lookup"><span data-stu-id="236b7-490">Attributes</span></span>| <span data-ttu-id="236b7-491">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="236b7-492">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-492">Object</span></span> | <span data-ttu-id="236b7-493">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-493">&lt;optional&gt;</span></span> | <span data-ttu-id="236b7-494">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="236b7-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="236b7-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="236b7-495">Boolean</span></span> |  <span data-ttu-id="236b7-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-496">&lt;optional&gt;</span></span> | <span data-ttu-id="236b7-p134">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="236b7-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="236b7-499">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-499">Object</span></span> |  <span data-ttu-id="236b7-500">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-500">&lt;optional&gt;</span></span> | <span data-ttu-id="236b7-501">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="236b7-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="236b7-502">fonction</span><span class="sxs-lookup"><span data-stu-id="236b7-502">function</span></span>||<span data-ttu-id="236b7-p135">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="236b7-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-505">Requirements</span></span>

|<span data-ttu-id="236b7-506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-506">Requirement</span></span>| <span data-ttu-id="236b7-507">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-509">1,5</span><span class="sxs-lookup"><span data-stu-id="236b7-509">1.5</span></span> |
|[<span data-ttu-id="236b7-510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-511">ReadItem</span></span>|
|[<span data-ttu-id="236b7-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-513">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-514">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="236b7-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="236b7-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="236b7-516">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="236b7-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="236b7-p136">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="236b7-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="236b7-p137">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="236b7-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="236b7-522">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="236b7-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="236b7-p138">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="236b7-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-525">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-525">Parameters</span></span>

|<span data-ttu-id="236b7-526">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-526">Name</span></span>| <span data-ttu-id="236b7-527">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-527">Type</span></span>| <span data-ttu-id="236b7-528">Attributs</span><span class="sxs-lookup"><span data-stu-id="236b7-528">Attributes</span></span>| <span data-ttu-id="236b7-529">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="236b7-530">fonction</span><span class="sxs-lookup"><span data-stu-id="236b7-530">function</span></span>||<span data-ttu-id="236b7-p139">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="236b7-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="236b7-533">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-533">Object</span></span>| <span data-ttu-id="236b7-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-534">&lt;optional&gt;</span></span>|<span data-ttu-id="236b7-535">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="236b7-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-536">Requirements</span></span>

|<span data-ttu-id="236b7-537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-537">Requirement</span></span>| <span data-ttu-id="236b7-538">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-539">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-540">1.3</span><span class="sxs-lookup"><span data-stu-id="236b7-540">1.3</span></span>|
|[<span data-ttu-id="236b7-541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-542">ReadItem</span></span>|
|[<span data-ttu-id="236b7-543">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-544">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-545">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-545">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="236b7-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="236b7-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="236b7-547">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="236b7-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="236b7-548">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="236b7-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-549">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-549">Parameters</span></span>

|<span data-ttu-id="236b7-550">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-550">Name</span></span>| <span data-ttu-id="236b7-551">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-551">Type</span></span>| <span data-ttu-id="236b7-552">Attributs</span><span class="sxs-lookup"><span data-stu-id="236b7-552">Attributes</span></span>| <span data-ttu-id="236b7-553">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="236b7-554">function</span><span class="sxs-lookup"><span data-stu-id="236b7-554">function</span></span>||<span data-ttu-id="236b7-555">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="236b7-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="236b7-556">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="236b7-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="236b7-557">Object</span><span class="sxs-lookup"><span data-stu-id="236b7-557">Object</span></span>| <span data-ttu-id="236b7-558">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-558">&lt;optional&gt;</span></span>|<span data-ttu-id="236b7-559">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="236b7-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-560">Requirements</span></span>

|<span data-ttu-id="236b7-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-561">Requirement</span></span>| <span data-ttu-id="236b7-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-564">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-564">1.0</span></span>|
|[<span data-ttu-id="236b7-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-566">ReadItem</span></span>|
|[<span data-ttu-id="236b7-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-568">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-569">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-569">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="236b7-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="236b7-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="236b7-571">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="236b7-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-572">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="236b7-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="236b7-573">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="236b7-573">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="236b7-574">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="236b7-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="236b7-575">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="236b7-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="236b7-576">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="236b7-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="236b7-577">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="236b7-577">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="236b7-578">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="236b7-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="236b7-579">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="236b7-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="236b7-p141">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="236b7-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="236b7-582">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="236b7-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="236b7-583">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="236b7-583">Version differences</span></span>

<span data-ttu-id="236b7-584">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="236b7-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="236b7-p142">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="236b7-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-588">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-588">Parameters</span></span>

|<span data-ttu-id="236b7-589">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-589">Name</span></span>| <span data-ttu-id="236b7-590">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-590">Type</span></span>| <span data-ttu-id="236b7-591">Attributs</span><span class="sxs-lookup"><span data-stu-id="236b7-591">Attributes</span></span>| <span data-ttu-id="236b7-592">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="236b7-593">String</span><span class="sxs-lookup"><span data-stu-id="236b7-593">String</span></span>||<span data-ttu-id="236b7-594">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="236b7-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="236b7-595">function</span><span class="sxs-lookup"><span data-stu-id="236b7-595">function</span></span>||<span data-ttu-id="236b7-596">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="236b7-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="236b7-597">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="236b7-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="236b7-598">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="236b7-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="236b7-599">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-599">Object</span></span>| <span data-ttu-id="236b7-600">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-600">&lt;optional&gt;</span></span>|<span data-ttu-id="236b7-601">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="236b7-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-602">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-602">Requirements</span></span>

|<span data-ttu-id="236b7-603">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-603">Requirement</span></span>| <span data-ttu-id="236b7-604">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-605">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-606">1.0</span><span class="sxs-lookup"><span data-stu-id="236b7-606">1.0</span></span>|
|[<span data-ttu-id="236b7-607">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="236b7-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="236b7-609">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-610">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="236b7-611">Exemple</span><span class="sxs-lookup"><span data-stu-id="236b7-611">Example</span></span>

<span data-ttu-id="236b7-612">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="236b7-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="236b7-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="236b7-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="236b7-614">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="236b7-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="236b7-615">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="236b7-615">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="236b7-616">Paramètres</span><span class="sxs-lookup"><span data-stu-id="236b7-616">Parameters</span></span>

| <span data-ttu-id="236b7-617">Nom</span><span class="sxs-lookup"><span data-stu-id="236b7-617">Name</span></span> | <span data-ttu-id="236b7-618">Type</span><span class="sxs-lookup"><span data-stu-id="236b7-618">Type</span></span> | <span data-ttu-id="236b7-619">Attributs</span><span class="sxs-lookup"><span data-stu-id="236b7-619">Attributes</span></span> | <span data-ttu-id="236b7-620">Description</span><span class="sxs-lookup"><span data-stu-id="236b7-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="236b7-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="236b7-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="236b7-622">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="236b7-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="236b7-623">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-623">Object</span></span> | <span data-ttu-id="236b7-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-624">&lt;optional&gt;</span></span> | <span data-ttu-id="236b7-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="236b7-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="236b7-626">Objet</span><span class="sxs-lookup"><span data-stu-id="236b7-626">Object</span></span> | <span data-ttu-id="236b7-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-627">&lt;optional&gt;</span></span> | <span data-ttu-id="236b7-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="236b7-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="236b7-629">fonction</span><span class="sxs-lookup"><span data-stu-id="236b7-629">function</span></span>| <span data-ttu-id="236b7-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="236b7-630">&lt;optional&gt;</span></span>|<span data-ttu-id="236b7-631">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="236b7-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="236b7-632">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="236b7-632">Requirements</span></span>

|<span data-ttu-id="236b7-633">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="236b7-633">Requirement</span></span>| <span data-ttu-id="236b7-634">Valeur</span><span class="sxs-lookup"><span data-stu-id="236b7-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="236b7-635">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="236b7-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="236b7-636">1,5</span><span class="sxs-lookup"><span data-stu-id="236b7-636">1.5</span></span> |
|[<span data-ttu-id="236b7-637">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="236b7-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="236b7-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="236b7-638">ReadItem</span></span> |
|[<span data-ttu-id="236b7-639">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="236b7-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="236b7-640">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="236b7-640">Compose or Read</span></span>|
