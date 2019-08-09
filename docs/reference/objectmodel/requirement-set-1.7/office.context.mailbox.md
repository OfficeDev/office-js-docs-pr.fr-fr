---
title: Office. Context. Mailbox-ensemble de conditions requises 1,7
description: ''
ms.date: 08/06/2019
localization_priority: Normal
ms.openlocfilehash: 88b99a541653138ea9457d417d767ce8aa516cea
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268382"
---
# <a name="mailbox"></a><span data-ttu-id="2fab2-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="2fab2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="2fab2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="2fab2-104">Permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="2fab2-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="2fab2-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-105">Requirements</span></span>

|<span data-ttu-id="2fab2-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-106">Requirement</span></span>| <span data-ttu-id="2fab2-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-109">1.0</span></span>|
|[<span data-ttu-id="2fab2-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="2fab2-111">Restricted</span></span>|
|[<span data-ttu-id="2fab2-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2fab2-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="2fab2-114">Members and methods</span></span>

| <span data-ttu-id="2fab2-115">Membre</span><span class="sxs-lookup"><span data-stu-id="2fab2-115">Member</span></span> | <span data-ttu-id="2fab2-116">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2fab2-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="2fab2-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="2fab2-118">Membre</span><span class="sxs-lookup"><span data-stu-id="2fab2-118">Member</span></span> |
| [<span data-ttu-id="2fab2-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="2fab2-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="2fab2-120">Membre</span><span class="sxs-lookup"><span data-stu-id="2fab2-120">Member</span></span> |
| [<span data-ttu-id="2fab2-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2fab2-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="2fab2-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-122">Method</span></span> |
| [<span data-ttu-id="2fab2-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="2fab2-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="2fab2-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-124">Method</span></span> |
| [<span data-ttu-id="2fab2-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2fab2-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="2fab2-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-126">Method</span></span> |
| [<span data-ttu-id="2fab2-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="2fab2-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="2fab2-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-128">Method</span></span> |
| [<span data-ttu-id="2fab2-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="2fab2-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="2fab2-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-130">Method</span></span> |
| [<span data-ttu-id="2fab2-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="2fab2-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="2fab2-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-132">Method</span></span> |
| [<span data-ttu-id="2fab2-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="2fab2-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="2fab2-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-134">Method</span></span> |
| [<span data-ttu-id="2fab2-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="2fab2-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="2fab2-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-136">Method</span></span> |
| [<span data-ttu-id="2fab2-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="2fab2-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="2fab2-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-138">Method</span></span> |
| [<span data-ttu-id="2fab2-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2fab2-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="2fab2-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-140">Method</span></span> |
| [<span data-ttu-id="2fab2-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2fab2-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="2fab2-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-142">Method</span></span> |
| [<span data-ttu-id="2fab2-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="2fab2-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="2fab2-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-144">Method</span></span> |
| [<span data-ttu-id="2fab2-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="2fab2-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="2fab2-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-146">Method</span></span> |
| [<span data-ttu-id="2fab2-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="2fab2-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="2fab2-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="2fab2-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="2fab2-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="2fab2-149">Namespaces</span></span>

<span data-ttu-id="2fab2-150">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2fab2-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="2fab2-151">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2fab2-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="2fab2-152">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2fab2-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="2fab2-153">Membres</span><span class="sxs-lookup"><span data-stu-id="2fab2-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="2fab2-154">ewsUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-154">ewsUrl: String</span></span>

<span data-ttu-id="2fab2-155">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="2fab2-155">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="2fab2-156">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="2fab2-156">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-157">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="2fab2-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2fab2-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="2fab2-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2fab2-160">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="2fab2-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="2fab2-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="2fab2-163">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-163">Type</span></span>

*   <span data-ttu-id="2fab2-164">String</span><span class="sxs-lookup"><span data-stu-id="2fab2-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2fab2-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-165">Requirements</span></span>

|<span data-ttu-id="2fab2-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-166">Requirement</span></span>| <span data-ttu-id="2fab2-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-169">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-169">1.0</span></span>|
|[<span data-ttu-id="2fab2-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-171">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="2fab2-174">restUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-174">restUrl: String</span></span>

<span data-ttu-id="2fab2-175">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="2fab2-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="2fab2-176">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2fab2-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="2fab2-177">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="2fab2-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="2fab2-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="2fab2-180">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-180">Type</span></span>

*   <span data-ttu-id="2fab2-181">String</span><span class="sxs-lookup"><span data-stu-id="2fab2-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2fab2-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-182">Requirements</span></span>

|<span data-ttu-id="2fab2-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-183">Requirement</span></span>| <span data-ttu-id="2fab2-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-186">1,5</span><span class="sxs-lookup"><span data-stu-id="2fab2-186">1.5</span></span> |
|[<span data-ttu-id="2fab2-187">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-188">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="2fab2-191">Méthodes</span><span class="sxs-lookup"><span data-stu-id="2fab2-191">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="2fab2-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2fab2-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="2fab2-193">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2fab2-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="2fab2-194">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2fab2-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-195">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-195">Parameters</span></span>

| <span data-ttu-id="2fab2-196">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-196">Name</span></span> | <span data-ttu-id="2fab2-197">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-197">Type</span></span> | <span data-ttu-id="2fab2-198">Attributs</span><span class="sxs-lookup"><span data-stu-id="2fab2-198">Attributes</span></span> | <span data-ttu-id="2fab2-199">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2fab2-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2fab2-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2fab2-201">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="2fab2-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="2fab2-202">Fonction</span><span class="sxs-lookup"><span data-stu-id="2fab2-202">Function</span></span> || <span data-ttu-id="2fab2-p105">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="2fab2-206">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-206">Object</span></span> | <span data-ttu-id="2fab2-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-207">&lt;optional&gt;</span></span> | <span data-ttu-id="2fab2-208">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2fab2-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2fab2-209">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-209">Object</span></span> | <span data-ttu-id="2fab2-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-210">&lt;optional&gt;</span></span> | <span data-ttu-id="2fab2-211">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2fab2-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2fab2-212">fonction</span><span class="sxs-lookup"><span data-stu-id="2fab2-212">function</span></span>| <span data-ttu-id="2fab2-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-213">&lt;optional&gt;</span></span>|<span data-ttu-id="2fab2-214">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2fab2-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-215">Requirements</span></span>

|<span data-ttu-id="2fab2-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-216">Requirement</span></span>| <span data-ttu-id="2fab2-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-219">1,5</span><span class="sxs-lookup"><span data-stu-id="2fab2-219">1.5</span></span> |
|[<span data-ttu-id="2fab2-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-221">ReadItem</span></span> |
|[<span data-ttu-id="2fab2-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-223">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="2fab2-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="2fab2-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="2fab2-226">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="2fab2-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-227">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="2fab2-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2fab2-p106">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-230">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-230">Parameters</span></span>

|<span data-ttu-id="2fab2-231">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-231">Name</span></span>| <span data-ttu-id="2fab2-232">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-232">Type</span></span>| <span data-ttu-id="2fab2-233">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2fab2-234">String</span><span class="sxs-lookup"><span data-stu-id="2fab2-234">String</span></span>|<span data-ttu-id="2fab2-235">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="2fab2-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="2fab2-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="2fab2-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="2fab2-237">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="2fab2-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-238">Requirements</span></span>

|<span data-ttu-id="2fab2-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-239">Requirement</span></span>| <span data-ttu-id="2fab2-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-242">1.3</span><span class="sxs-lookup"><span data-stu-id="2fab2-242">1.3</span></span>|
|[<span data-ttu-id="2fab2-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-244">Restreinte</span><span class="sxs-lookup"><span data-stu-id="2fab2-244">Restricted</span></span>|
|[<span data-ttu-id="2fab2-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2fab2-247">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2fab2-247">Returns:</span></span>

<span data-ttu-id="2fab2-248">Type : String</span><span class="sxs-lookup"><span data-stu-id="2fab2-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2fab2-249">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="2fab2-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="2fab2-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="2fab2-251">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="2fab2-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="2fab2-252">Une application de messagerie pour Outlook sur un ordinateur de bureau ou sur le Web peut utiliser différents fuseaux horaires pour les dates et les heures.</span><span class="sxs-lookup"><span data-stu-id="2fab2-252">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="2fab2-253">Outlook sur un ordinateur de bureau utilise le fuseau horaire de l’ordinateur client; Outlook sur le Web utilise le fuseau horaire défini dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="2fab2-253">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="2fab2-254">Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2fab2-254">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="2fab2-255">Si l’application de messagerie est en cours d’exécution dans Outlook sur un `convertToLocalClientTime` client de bureau, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="2fab2-255">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="2fab2-256">Si l’application de messagerie est en cours d’exécution dans Outlook sur `convertToLocalClientTime` le Web, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire spécifié dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="2fab2-256">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-257">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-257">Parameters</span></span>

|<span data-ttu-id="2fab2-258">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-258">Name</span></span>| <span data-ttu-id="2fab2-259">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-259">Type</span></span>| <span data-ttu-id="2fab2-260">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="2fab2-261">Date</span><span class="sxs-lookup"><span data-stu-id="2fab2-261">Date</span></span>|<span data-ttu-id="2fab2-262">Objet Date</span><span class="sxs-lookup"><span data-stu-id="2fab2-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-263">Requirements</span></span>

|<span data-ttu-id="2fab2-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-264">Requirement</span></span>| <span data-ttu-id="2fab2-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-267">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-267">1.0</span></span>|
|[<span data-ttu-id="2fab2-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-269">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-271">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2fab2-272">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2fab2-272">Returns:</span></span>

<span data-ttu-id="2fab2-273">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="2fab2-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="2fab2-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="2fab2-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="2fab2-275">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="2fab2-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-276">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="2fab2-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2fab2-p109">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-279">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-279">Parameters</span></span>

|<span data-ttu-id="2fab2-280">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-280">Name</span></span>| <span data-ttu-id="2fab2-281">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-281">Type</span></span>| <span data-ttu-id="2fab2-282">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2fab2-283">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-283">String</span></span>|<span data-ttu-id="2fab2-284">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="2fab2-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="2fab2-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="2fab2-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="2fab2-286">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="2fab2-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-287">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-287">Requirements</span></span>

|<span data-ttu-id="2fab2-288">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-288">Requirement</span></span>| <span data-ttu-id="2fab2-289">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-290">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-291">1.3</span><span class="sxs-lookup"><span data-stu-id="2fab2-291">1.3</span></span>|
|[<span data-ttu-id="2fab2-292">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-293">Restreinte</span><span class="sxs-lookup"><span data-stu-id="2fab2-293">Restricted</span></span>|
|[<span data-ttu-id="2fab2-294">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-295">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2fab2-296">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2fab2-296">Returns:</span></span>

<span data-ttu-id="2fab2-297">Type : String</span><span class="sxs-lookup"><span data-stu-id="2fab2-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="2fab2-298">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="2fab2-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="2fab2-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="2fab2-300">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="2fab2-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="2fab2-301">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="2fab2-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-302">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-302">Parameters</span></span>

|<span data-ttu-id="2fab2-303">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-303">Name</span></span>| <span data-ttu-id="2fab2-304">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-304">Type</span></span>| <span data-ttu-id="2fab2-305">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="2fab2-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="2fab2-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="2fab2-307">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="2fab2-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-308">Requirements</span></span>

|<span data-ttu-id="2fab2-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-309">Requirement</span></span>| <span data-ttu-id="2fab2-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-312">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-312">1.0</span></span>|
|[<span data-ttu-id="2fab2-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-314">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-316">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="2fab2-317">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="2fab2-317">Returns:</span></span>

<span data-ttu-id="2fab2-318">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="2fab2-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="2fab2-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="2fab2-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="2fab2-320">Date</span><span class="sxs-lookup"><span data-stu-id="2fab2-320">Date</span></span></dd>

</dl>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="2fab2-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2fab2-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="2fab2-322">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="2fab2-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-323">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="2fab2-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2fab2-324">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="2fab2-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2fab2-325">Dans Outlook sur Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série.</span><span class="sxs-lookup"><span data-stu-id="2fab2-325">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="2fab2-326">En effet, dans Outlook sur Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="2fab2-326">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="2fab2-327">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32KO nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="2fab2-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="2fab2-328">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="2fab2-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-329">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-329">Parameters</span></span>

|<span data-ttu-id="2fab2-330">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-330">Name</span></span>| <span data-ttu-id="2fab2-331">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-331">Type</span></span>| <span data-ttu-id="2fab2-332">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2fab2-333">String</span><span class="sxs-lookup"><span data-stu-id="2fab2-333">String</span></span>|<span data-ttu-id="2fab2-334">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="2fab2-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-335">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-335">Requirements</span></span>

|<span data-ttu-id="2fab2-336">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-336">Requirement</span></span>| <span data-ttu-id="2fab2-337">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-338">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-339">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-339">1.0</span></span>|
|[<span data-ttu-id="2fab2-340">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-341">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-342">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-343">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-344">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="2fab2-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="2fab2-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="2fab2-346">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="2fab2-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-347">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="2fab2-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2fab2-348">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="2fab2-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="2fab2-349">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32 Ko nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="2fab2-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="2fab2-350">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="2fab2-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="2fab2-p111">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-353">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-353">Parameters</span></span>

|<span data-ttu-id="2fab2-354">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-354">Name</span></span>| <span data-ttu-id="2fab2-355">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-355">Type</span></span>| <span data-ttu-id="2fab2-356">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="2fab2-357">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-357">String</span></span>|<span data-ttu-id="2fab2-358">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="2fab2-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-359">Requirements</span></span>

|<span data-ttu-id="2fab2-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-360">Requirement</span></span>| <span data-ttu-id="2fab2-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-363">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-363">1.0</span></span>|
|[<span data-ttu-id="2fab2-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-365">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-367">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-368">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="2fab2-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="2fab2-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="2fab2-370">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="2fab2-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-371">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="2fab2-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="2fab2-p112">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2fab2-374">Dans Outlook sur le Web et les appareils mobiles, cette méthode affiche toujours un formulaire avec un champ participants.</span><span class="sxs-lookup"><span data-stu-id="2fab2-374">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="2fab2-375">Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="2fab2-375">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="2fab2-376">Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="2fab2-376">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="2fab2-p114">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="2fab2-379">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="2fab2-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-380">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-381">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="2fab2-381">All parameters are optional.</span></span>

|<span data-ttu-id="2fab2-382">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-382">Name</span></span>| <span data-ttu-id="2fab2-383">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-383">Type</span></span>| <span data-ttu-id="2fab2-384">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2fab2-385">Object</span><span class="sxs-lookup"><span data-stu-id="2fab2-385">Object</span></span> | <span data-ttu-id="2fab2-386">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2fab2-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="2fab2-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="2fab2-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="2fab2-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="2fab2-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="2fab2-393">Date</span><span class="sxs-lookup"><span data-stu-id="2fab2-393">Date</span></span> | <span data-ttu-id="2fab2-394">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2fab2-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="2fab2-395">Date</span><span class="sxs-lookup"><span data-stu-id="2fab2-395">Date</span></span> | <span data-ttu-id="2fab2-396">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="2fab2-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="2fab2-397">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-397">String</span></span> | <span data-ttu-id="2fab2-p117">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="2fab2-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="2fab2-p118">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2fab2-403">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-403">String</span></span> | <span data-ttu-id="2fab2-p119">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="2fab2-406">String</span><span class="sxs-lookup"><span data-stu-id="2fab2-406">String</span></span> | <span data-ttu-id="2fab2-p120">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2fab2-409">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-409">Requirements</span></span>

|<span data-ttu-id="2fab2-410">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-410">Requirement</span></span>| <span data-ttu-id="2fab2-411">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-412">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-413">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-413">1.0</span></span>|
|[<span data-ttu-id="2fab2-414">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-415">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-416">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-417">Lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-418">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="2fab2-419">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="2fab2-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="2fab2-420">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="2fab2-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="2fab2-421">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="2fab2-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="2fab2-422">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="2fab2-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="2fab2-423">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="2fab2-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-424">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-425">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="2fab2-425">All parameters are optional.</span></span>

|<span data-ttu-id="2fab2-426">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-426">Name</span></span>| <span data-ttu-id="2fab2-427">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-427">Type</span></span>| <span data-ttu-id="2fab2-428">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="2fab2-429">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-429">Object</span></span> | <span data-ttu-id="2fab2-430">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="2fab2-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="2fab2-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="2fab2-432">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="2fab2-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="2fab2-433">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="2fab2-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="2fab2-435">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="2fab2-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="2fab2-436">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="2fab2-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="2fab2-438">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="2fab2-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="2fab2-439">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="2fab2-440">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-440">String</span></span> | <span data-ttu-id="2fab2-441">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="2fab2-441">A string containing the subject of the message.</span></span> <span data-ttu-id="2fab2-442">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="2fab2-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="2fab2-443">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-443">String</span></span> | <span data-ttu-id="2fab2-444">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="2fab2-444">The HTML body of the message.</span></span> <span data-ttu-id="2fab2-445">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="2fab2-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="2fab2-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="2fab2-447">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="2fab2-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="2fab2-448">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-448">String</span></span> | <span data-ttu-id="2fab2-p127">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="2fab2-451">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-451">String</span></span> | <span data-ttu-id="2fab2-452">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="2fab2-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="2fab2-453">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-453">String</span></span> | <span data-ttu-id="2fab2-p128">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="2fab2-456">Booléen</span><span class="sxs-lookup"><span data-stu-id="2fab2-456">Boolean</span></span> | <span data-ttu-id="2fab2-p129">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="2fab2-459">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fab2-459">String</span></span> | <span data-ttu-id="2fab2-460">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="2fab2-461">ID d’élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="2fab2-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="2fab2-462">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="2fab2-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="2fab2-463">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-463">Requirements</span></span>

|<span data-ttu-id="2fab2-464">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-464">Requirement</span></span>| <span data-ttu-id="2fab2-465">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-466">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-467">1.6</span><span class="sxs-lookup"><span data-stu-id="2fab2-467">1.6</span></span> |
|[<span data-ttu-id="2fab2-468">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-469">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-470">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-471">Lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-472">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-472">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="2fab2-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="2fab2-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="2fab2-474">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="2fab2-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="2fab2-p131">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-477">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="2fab2-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="2fab2-478">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="2fab2-478">**REST Tokens**</span></span>

<span data-ttu-id="2fab2-p132">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="2fab2-482">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="2fab2-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="2fab2-483">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="2fab2-483">**EWS Tokens**</span></span>

<span data-ttu-id="2fab2-p133">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="2fab2-486">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="2fab2-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-487">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-487">Parameters</span></span>

|<span data-ttu-id="2fab2-488">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-488">Name</span></span>| <span data-ttu-id="2fab2-489">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-489">Type</span></span>| <span data-ttu-id="2fab2-490">Attributs</span><span class="sxs-lookup"><span data-stu-id="2fab2-490">Attributes</span></span>| <span data-ttu-id="2fab2-491">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="2fab2-492">Object</span><span class="sxs-lookup"><span data-stu-id="2fab2-492">Object</span></span> | <span data-ttu-id="2fab2-493">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-493">&lt;optional&gt;</span></span> | <span data-ttu-id="2fab2-494">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2fab2-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="2fab2-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="2fab2-495">Boolean</span></span> |  <span data-ttu-id="2fab2-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-496">&lt;optional&gt;</span></span> | <span data-ttu-id="2fab2-p134">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2fab2-499">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-499">Object</span></span> |  <span data-ttu-id="2fab2-500">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-500">&lt;optional&gt;</span></span> | <span data-ttu-id="2fab2-501">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2fab2-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="2fab2-502">fonction</span><span class="sxs-lookup"><span data-stu-id="2fab2-502">function</span></span>||<span data-ttu-id="2fab2-503">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2fab2-503">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2fab2-504">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-504">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="2fab2-505">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="2fab2-505">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2fab2-506">Erreurs</span><span class="sxs-lookup"><span data-stu-id="2fab2-506">Errors</span></span>

|<span data-ttu-id="2fab2-507">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="2fab2-507">Error code</span></span>|<span data-ttu-id="2fab2-508">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-508">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="2fab2-509">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="2fab2-509">The request has failed.</span></span> <span data-ttu-id="2fab2-510">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="2fab2-510">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="2fab2-511">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="2fab2-511">The Exchange server returned an error.</span></span> <span data-ttu-id="2fab2-512">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="2fab2-512">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="2fab2-513">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="2fab2-513">The user is no longer connected to the network.</span></span> <span data-ttu-id="2fab2-514">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="2fab2-514">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-515">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-515">Requirements</span></span>

|<span data-ttu-id="2fab2-516">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-516">Requirement</span></span>| <span data-ttu-id="2fab2-517">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-517">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-518">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-518">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-519">1,5</span><span class="sxs-lookup"><span data-stu-id="2fab2-519">1.5</span></span> |
|[<span data-ttu-id="2fab2-520">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-520">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-521">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-521">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-522">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-522">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-523">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-523">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-524">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-524">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="2fab2-525">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2fab2-525">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2fab2-526">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="2fab2-526">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="2fab2-p138">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p138">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="2fab2-p139">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="2fab2-p139">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="2fab2-532">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="2fab2-532">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="2fab2-p140">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p140">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-535">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-535">Parameters</span></span>

|<span data-ttu-id="2fab2-536">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-536">Name</span></span>| <span data-ttu-id="2fab2-537">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-537">Type</span></span>| <span data-ttu-id="2fab2-538">Attributs</span><span class="sxs-lookup"><span data-stu-id="2fab2-538">Attributes</span></span>| <span data-ttu-id="2fab2-539">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-539">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2fab2-540">function</span><span class="sxs-lookup"><span data-stu-id="2fab2-540">function</span></span>||<span data-ttu-id="2fab2-541">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2fab2-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2fab2-542">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-542">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="2fab2-543">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="2fab2-543">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="2fab2-544">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-544">Object</span></span>| <span data-ttu-id="2fab2-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-545">&lt;optional&gt;</span></span>|<span data-ttu-id="2fab2-546">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2fab2-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2fab2-547">Erreurs</span><span class="sxs-lookup"><span data-stu-id="2fab2-547">Errors</span></span>

|<span data-ttu-id="2fab2-548">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="2fab2-548">Error code</span></span>|<span data-ttu-id="2fab2-549">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-549">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="2fab2-550">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="2fab2-550">The request has failed.</span></span> <span data-ttu-id="2fab2-551">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="2fab2-551">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="2fab2-552">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="2fab2-552">The Exchange server returned an error.</span></span> <span data-ttu-id="2fab2-553">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="2fab2-553">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="2fab2-554">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="2fab2-554">The user is no longer connected to the network.</span></span> <span data-ttu-id="2fab2-555">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="2fab2-555">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-556">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-556">Requirements</span></span>

|<span data-ttu-id="2fab2-557">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-557">Requirement</span></span>| <span data-ttu-id="2fab2-558">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-558">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-559">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-559">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-560">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-560">1.0</span></span>|
|[<span data-ttu-id="2fab2-561">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-561">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-562">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-562">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-563">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-563">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-564">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-564">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-565">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-565">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="2fab2-566">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2fab2-566">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="2fab2-567">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="2fab2-567">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="2fab2-568">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="2fab2-568">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-569">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-569">Parameters</span></span>

|<span data-ttu-id="2fab2-570">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-570">Name</span></span>| <span data-ttu-id="2fab2-571">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-571">Type</span></span>| <span data-ttu-id="2fab2-572">Attributs</span><span class="sxs-lookup"><span data-stu-id="2fab2-572">Attributes</span></span>| <span data-ttu-id="2fab2-573">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-573">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="2fab2-574">function</span><span class="sxs-lookup"><span data-stu-id="2fab2-574">function</span></span>||<span data-ttu-id="2fab2-575">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2fab2-575">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2fab2-576">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-576">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="2fab2-577">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="2fab2-577">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="2fab2-578">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-578">Object</span></span>| <span data-ttu-id="2fab2-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-579">&lt;optional&gt;</span></span>|<span data-ttu-id="2fab2-580">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2fab2-580">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="2fab2-581">Erreurs</span><span class="sxs-lookup"><span data-stu-id="2fab2-581">Errors</span></span>

|<span data-ttu-id="2fab2-582">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="2fab2-582">Error code</span></span>|<span data-ttu-id="2fab2-583">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-583">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="2fab2-584">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="2fab2-584">The request has failed.</span></span> <span data-ttu-id="2fab2-585">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="2fab2-585">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="2fab2-586">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="2fab2-586">The Exchange server returned an error.</span></span> <span data-ttu-id="2fab2-587">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="2fab2-587">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="2fab2-588">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="2fab2-588">The user is no longer connected to the network.</span></span> <span data-ttu-id="2fab2-589">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="2fab2-589">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-590">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-590">Requirements</span></span>

|<span data-ttu-id="2fab2-591">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-591">Requirement</span></span>| <span data-ttu-id="2fab2-592">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-592">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-593">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-593">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-594">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-594">1.0</span></span>|
|[<span data-ttu-id="2fab2-595">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-595">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-596">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-596">ReadItem</span></span>|
|[<span data-ttu-id="2fab2-597">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-597">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-598">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-598">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-599">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-599">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="2fab2-600">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="2fab2-600">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="2fab2-601">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2fab2-601">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-602">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="2fab2-602">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="2fab2-603">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="2fab2-603">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="2fab2-604">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="2fab2-604">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="2fab2-605">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2fab2-605">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="2fab2-606">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="2fab2-606">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="2fab2-607">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="2fab2-607">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="2fab2-608">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-608">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="2fab2-609">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="2fab2-609">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="2fab2-p148">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="2fab2-p148">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="2fab2-612">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="2fab2-612">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="2fab2-613">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="2fab2-613">Version differences</span></span>

<span data-ttu-id="2fab2-614">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-614">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="2fab2-p149">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="2fab2-p149">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-618">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-618">Parameters</span></span>

|<span data-ttu-id="2fab2-619">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-619">Name</span></span>| <span data-ttu-id="2fab2-620">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-620">Type</span></span>| <span data-ttu-id="2fab2-621">Attributs</span><span class="sxs-lookup"><span data-stu-id="2fab2-621">Attributes</span></span>| <span data-ttu-id="2fab2-622">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-622">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="2fab2-623">String</span><span class="sxs-lookup"><span data-stu-id="2fab2-623">String</span></span>||<span data-ttu-id="2fab2-624">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="2fab2-624">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="2fab2-625">function</span><span class="sxs-lookup"><span data-stu-id="2fab2-625">function</span></span>||<span data-ttu-id="2fab2-626">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2fab2-626">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="2fab2-627">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="2fab2-627">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="2fab2-628">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="2fab2-628">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="2fab2-629">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-629">Object</span></span>| <span data-ttu-id="2fab2-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-630">&lt;optional&gt;</span></span>|<span data-ttu-id="2fab2-631">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2fab2-631">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-632">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-632">Requirements</span></span>

|<span data-ttu-id="2fab2-633">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-633">Requirement</span></span>| <span data-ttu-id="2fab2-634">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-635">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-636">1.0</span><span class="sxs-lookup"><span data-stu-id="2fab2-636">1.0</span></span>|
|[<span data-ttu-id="2fab2-637">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-638">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="2fab2-638">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="2fab2-639">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-640">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-640">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2fab2-641">Exemple</span><span class="sxs-lookup"><span data-stu-id="2fab2-641">Example</span></span>

<span data-ttu-id="2fab2-642">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="2fab2-642">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="2fab2-643">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="2fab2-643">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="2fab2-644">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2fab2-644">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="2fab2-645">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="2fab2-645">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="2fab2-646">Paramètres</span><span class="sxs-lookup"><span data-stu-id="2fab2-646">Parameters</span></span>

| <span data-ttu-id="2fab2-647">Nom</span><span class="sxs-lookup"><span data-stu-id="2fab2-647">Name</span></span> | <span data-ttu-id="2fab2-648">Type</span><span class="sxs-lookup"><span data-stu-id="2fab2-648">Type</span></span> | <span data-ttu-id="2fab2-649">Attributs</span><span class="sxs-lookup"><span data-stu-id="2fab2-649">Attributes</span></span> | <span data-ttu-id="2fab2-650">Description</span><span class="sxs-lookup"><span data-stu-id="2fab2-650">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="2fab2-651">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="2fab2-651">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="2fab2-652">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="2fab2-652">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="2fab2-653">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-653">Object</span></span> | <span data-ttu-id="2fab2-654">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-654">&lt;optional&gt;</span></span> | <span data-ttu-id="2fab2-655">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="2fab2-655">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="2fab2-656">Objet</span><span class="sxs-lookup"><span data-stu-id="2fab2-656">Object</span></span> | <span data-ttu-id="2fab2-657">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-657">&lt;optional&gt;</span></span> | <span data-ttu-id="2fab2-658">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="2fab2-658">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="2fab2-659">fonction</span><span class="sxs-lookup"><span data-stu-id="2fab2-659">function</span></span>| <span data-ttu-id="2fab2-660">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="2fab2-660">&lt;optional&gt;</span></span>|<span data-ttu-id="2fab2-661">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="2fab2-661">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fab2-662">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fab2-662">Requirements</span></span>

|<span data-ttu-id="2fab2-663">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fab2-663">Requirement</span></span>| <span data-ttu-id="2fab2-664">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fab2-664">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fab2-665">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fab2-665">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2fab2-666">1,5</span><span class="sxs-lookup"><span data-stu-id="2fab2-666">1.5</span></span> |
|[<span data-ttu-id="2fab2-667">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2fab2-667">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2fab2-668">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2fab2-668">ReadItem</span></span> |
|[<span data-ttu-id="2fab2-669">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fab2-669">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fab2-670">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fab2-670">Compose or Read</span></span>|
