---
title: Office. Context. Mailbox-ensemble de conditions requises 1,6
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 09c3930daf6f26edbc38b01f515ee5b1830ce802
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629692"
---
# <a name="mailbox"></a><span data-ttu-id="dd70d-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="dd70d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="dd70d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="dd70d-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="dd70d-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd70d-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-105">Requirements</span></span>

|<span data-ttu-id="dd70d-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-106">Requirement</span></span>| <span data-ttu-id="dd70d-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-109">1.0</span></span>|
|[<span data-ttu-id="dd70d-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dd70d-111">Restricted</span></span>|
|[<span data-ttu-id="dd70d-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="dd70d-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="dd70d-114">Members and methods</span></span>

| <span data-ttu-id="dd70d-115">Membre</span><span class="sxs-lookup"><span data-stu-id="dd70d-115">Member</span></span> | <span data-ttu-id="dd70d-116">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="dd70d-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="dd70d-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="dd70d-118">Membre</span><span class="sxs-lookup"><span data-stu-id="dd70d-118">Member</span></span> |
| [<span data-ttu-id="dd70d-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="dd70d-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="dd70d-120">Membre</span><span class="sxs-lookup"><span data-stu-id="dd70d-120">Member</span></span> |
| [<span data-ttu-id="dd70d-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="dd70d-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="dd70d-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-122">Method</span></span> |
| [<span data-ttu-id="dd70d-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="dd70d-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="dd70d-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-124">Method</span></span> |
| [<span data-ttu-id="dd70d-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="dd70d-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="dd70d-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-126">Method</span></span> |
| [<span data-ttu-id="dd70d-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="dd70d-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="dd70d-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-128">Method</span></span> |
| [<span data-ttu-id="dd70d-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="dd70d-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="dd70d-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-130">Method</span></span> |
| [<span data-ttu-id="dd70d-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dd70d-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="dd70d-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-132">Method</span></span> |
| [<span data-ttu-id="dd70d-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="dd70d-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="dd70d-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-134">Method</span></span> |
| [<span data-ttu-id="dd70d-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="dd70d-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="dd70d-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-136">Method</span></span> |
| [<span data-ttu-id="dd70d-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="dd70d-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="dd70d-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-138">Method</span></span> |
| [<span data-ttu-id="dd70d-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dd70d-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="dd70d-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-140">Method</span></span> |
| [<span data-ttu-id="dd70d-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dd70d-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="dd70d-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-142">Method</span></span> |
| [<span data-ttu-id="dd70d-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="dd70d-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="dd70d-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-144">Method</span></span> |
| [<span data-ttu-id="dd70d-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="dd70d-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="dd70d-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-146">Method</span></span> |
| [<span data-ttu-id="dd70d-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="dd70d-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="dd70d-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="dd70d-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="dd70d-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="dd70d-149">Namespaces</span></span>

<span data-ttu-id="dd70d-150">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="dd70d-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="dd70d-151">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="dd70d-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="dd70d-152">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="dd70d-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="dd70d-153">Members</span><span class="sxs-lookup"><span data-stu-id="dd70d-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="dd70d-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="dd70d-154">ewsUrl: String</span></span>

<span data-ttu-id="dd70d-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-157">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dd70d-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd70d-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dd70d-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="dd70d-160">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="dd70d-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="dd70d-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="dd70d-163">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-163">Type</span></span>

*   <span data-ttu-id="dd70d-164">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd70d-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-165">Requirements</span></span>

|<span data-ttu-id="dd70d-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-166">Requirement</span></span>| <span data-ttu-id="dd70d-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-169">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-169">1.0</span></span>|
|[<span data-ttu-id="dd70d-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-171">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="dd70d-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="dd70d-174">restUrl: String</span></span>

<span data-ttu-id="dd70d-175">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="dd70d-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="dd70d-176">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dd70d-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="dd70d-177">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-177">Type</span></span>

*   <span data-ttu-id="dd70d-178">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd70d-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-179">Requirements</span></span>

|<span data-ttu-id="dd70d-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-180">Requirement</span></span>| <span data-ttu-id="dd70d-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-183">1,5</span><span class="sxs-lookup"><span data-stu-id="dd70d-183">1.5</span></span> |
|[<span data-ttu-id="dd70d-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-185">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="dd70d-188">Méthodes</span><span class="sxs-lookup"><span data-stu-id="dd70d-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="dd70d-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd70d-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="dd70d-190">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="dd70d-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="dd70d-191">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé quand l’utilisateur sélectionne un nouvel élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-191">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="dd70d-192">Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’IU du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="dd70d-192">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-193">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-193">Parameters</span></span>

| <span data-ttu-id="dd70d-194">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-194">Name</span></span> | <span data-ttu-id="dd70d-195">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-195">Type</span></span> | <span data-ttu-id="dd70d-196">Attributs</span><span class="sxs-lookup"><span data-stu-id="dd70d-196">Attributes</span></span> | <span data-ttu-id="dd70d-197">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="dd70d-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="dd70d-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="dd70d-199">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="dd70d-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="dd70d-200">Fonction</span><span class="sxs-lookup"><span data-stu-id="dd70d-200">Function</span></span> || <span data-ttu-id="dd70d-p105">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="dd70d-204">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-204">Object</span></span> | <span data-ttu-id="dd70d-205">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-205">&lt;optional&gt;</span></span> | <span data-ttu-id="dd70d-206">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="dd70d-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dd70d-207">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-207">Object</span></span> | <span data-ttu-id="dd70d-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-208">&lt;optional&gt;</span></span> | <span data-ttu-id="dd70d-209">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="dd70d-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="dd70d-210">fonction</span><span class="sxs-lookup"><span data-stu-id="dd70d-210">function</span></span>| <span data-ttu-id="dd70d-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-211">&lt;optional&gt;</span></span>|<span data-ttu-id="dd70d-212">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd70d-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-213">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-213">Requirements</span></span>

|<span data-ttu-id="dd70d-214">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-214">Requirement</span></span>| <span data-ttu-id="dd70d-215">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-216">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-217">1,5</span><span class="sxs-lookup"><span data-stu-id="dd70d-217">1.5</span></span> |
|[<span data-ttu-id="dd70d-218">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-218">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-219">ReadItem</span></span> |
|[<span data-ttu-id="dd70d-220">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-220">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-221">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-222">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-222">Example</span></span>

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
};
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="dd70d-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="dd70d-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="dd70d-224">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="dd70d-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-225">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dd70d-225">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd70d-p106">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-228">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-228">Parameters</span></span>

|<span data-ttu-id="dd70d-229">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-229">Name</span></span>| <span data-ttu-id="dd70d-230">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-230">Type</span></span>| <span data-ttu-id="dd70d-231">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dd70d-232">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-232">String</span></span>|<span data-ttu-id="dd70d-233">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="dd70d-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="dd70d-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="dd70d-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="dd70d-235">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-236">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-236">Requirements</span></span>

|<span data-ttu-id="dd70d-237">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-237">Requirement</span></span>| <span data-ttu-id="dd70d-238">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-239">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-240">1.3</span><span class="sxs-lookup"><span data-stu-id="dd70d-240">1.3</span></span>|
|[<span data-ttu-id="dd70d-241">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-242">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dd70d-242">Restricted</span></span>|
|[<span data-ttu-id="dd70d-243">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-244">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-244">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd70d-245">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dd70d-245">Returns:</span></span>

<span data-ttu-id="dd70d-246">Type : String</span><span class="sxs-lookup"><span data-stu-id="dd70d-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="dd70d-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-247">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="dd70d-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="dd70d-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="dd70d-249">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="dd70d-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="dd70d-p107">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="dd70d-p108">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-255">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dd70d-255">Parameters</span></span>

|<span data-ttu-id="dd70d-256">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-256">Name</span></span>| <span data-ttu-id="dd70d-257">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-257">Type</span></span>| <span data-ttu-id="dd70d-258">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="dd70d-259">Date</span><span class="sxs-lookup"><span data-stu-id="dd70d-259">Date</span></span>|<span data-ttu-id="dd70d-260">Objet Date</span><span class="sxs-lookup"><span data-stu-id="dd70d-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-261">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-261">Requirements</span></span>

|<span data-ttu-id="dd70d-262">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-262">Requirement</span></span>| <span data-ttu-id="dd70d-263">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-264">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-265">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-265">1.0</span></span>|
|[<span data-ttu-id="dd70d-266">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-266">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-267">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-268">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-268">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-269">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-269">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd70d-270">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dd70d-270">Returns:</span></span>

<span data-ttu-id="dd70d-271">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="dd70d-271">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="dd70d-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="dd70d-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="dd70d-273">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="dd70d-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-274">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dd70d-274">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd70d-p109">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-277">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-277">Parameters</span></span>

|<span data-ttu-id="dd70d-278">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-278">Name</span></span>| <span data-ttu-id="dd70d-279">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-279">Type</span></span>| <span data-ttu-id="dd70d-280">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dd70d-281">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-281">String</span></span>|<span data-ttu-id="dd70d-282">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="dd70d-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="dd70d-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="dd70d-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="dd70d-284">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="dd70d-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-285">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-285">Requirements</span></span>

|<span data-ttu-id="dd70d-286">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-286">Requirement</span></span>| <span data-ttu-id="dd70d-287">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-288">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-289">1.3</span><span class="sxs-lookup"><span data-stu-id="dd70d-289">1.3</span></span>|
|[<span data-ttu-id="dd70d-290">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-291">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dd70d-291">Restricted</span></span>|
|[<span data-ttu-id="dd70d-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-293">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-293">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd70d-294">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dd70d-294">Returns:</span></span>

<span data-ttu-id="dd70d-295">Type : String</span><span class="sxs-lookup"><span data-stu-id="dd70d-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="dd70d-296">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-296">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="dd70d-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="dd70d-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="dd70d-298">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="dd70d-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="dd70d-299">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="dd70d-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-300">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-300">Parameters</span></span>

|<span data-ttu-id="dd70d-301">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-301">Name</span></span>| <span data-ttu-id="dd70d-302">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-302">Type</span></span>| <span data-ttu-id="dd70d-303">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="dd70d-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="dd70d-304">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="dd70d-305">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="dd70d-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-306">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-306">Requirements</span></span>

|<span data-ttu-id="dd70d-307">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-307">Requirement</span></span>| <span data-ttu-id="dd70d-308">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-309">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-310">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-310">1.0</span></span>|
|[<span data-ttu-id="dd70d-311">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-312">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-313">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-314">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-314">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="dd70d-315">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="dd70d-315">Returns:</span></span>

<span data-ttu-id="dd70d-316">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="dd70d-316">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="dd70d-317">Type : Date</span><span class="sxs-lookup"><span data-stu-id="dd70d-317">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="dd70d-318">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-318">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="dd70d-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dd70d-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="dd70d-320">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="dd70d-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-321">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dd70d-321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd70d-322">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="dd70d-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dd70d-p110">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="dd70d-325">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-325">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="dd70d-326">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="dd70d-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-327">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-327">Parameters</span></span>

|<span data-ttu-id="dd70d-328">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-328">Name</span></span>| <span data-ttu-id="dd70d-329">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-329">Type</span></span>| <span data-ttu-id="dd70d-330">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dd70d-331">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-331">String</span></span>|<span data-ttu-id="dd70d-332">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="dd70d-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-333">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-333">Requirements</span></span>

|<span data-ttu-id="dd70d-334">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-334">Requirement</span></span>| <span data-ttu-id="dd70d-335">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-336">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-337">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-337">1.0</span></span>|
|[<span data-ttu-id="dd70d-338">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-339">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-340">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-341">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-341">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-342">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-342">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="dd70d-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="dd70d-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="dd70d-344">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="dd70d-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-345">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dd70d-345">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd70d-346">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="dd70d-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="dd70d-347">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-347">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="dd70d-348">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="dd70d-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="dd70d-p111">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-351">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dd70d-351">Parameters</span></span>

|<span data-ttu-id="dd70d-352">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-352">Name</span></span>| <span data-ttu-id="dd70d-353">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-353">Type</span></span>| <span data-ttu-id="dd70d-354">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="dd70d-355">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-355">String</span></span>|<span data-ttu-id="dd70d-356">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="dd70d-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-357">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-357">Requirements</span></span>

|<span data-ttu-id="dd70d-358">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-358">Requirement</span></span>| <span data-ttu-id="dd70d-359">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-360">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-361">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-361">1.0</span></span>|
|[<span data-ttu-id="dd70d-362">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-363">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-364">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-365">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-365">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-366">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-366">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="dd70d-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="dd70d-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="dd70d-368">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="dd70d-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-369">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="dd70d-369">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="dd70d-p112">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="dd70d-p113">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="dd70d-p114">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="dd70d-377">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="dd70d-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-378">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dd70d-378">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-379">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="dd70d-379">All parameters are optional.</span></span>

|<span data-ttu-id="dd70d-380">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-380">Name</span></span>| <span data-ttu-id="dd70d-381">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-381">Type</span></span>| <span data-ttu-id="dd70d-382">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="dd70d-383">Object</span><span class="sxs-lookup"><span data-stu-id="dd70d-383">Object</span></span> | <span data-ttu-id="dd70d-384">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dd70d-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="dd70d-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="dd70d-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="dd70d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="dd70d-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="dd70d-391">Date</span><span class="sxs-lookup"><span data-stu-id="dd70d-391">Date</span></span> | <span data-ttu-id="dd70d-392">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dd70d-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="dd70d-393">Date</span><span class="sxs-lookup"><span data-stu-id="dd70d-393">Date</span></span> | <span data-ttu-id="dd70d-394">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="dd70d-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="dd70d-395">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dd70d-395">String</span></span> | <span data-ttu-id="dd70d-p117">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="dd70d-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="dd70d-p118">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="dd70d-401">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-401">String</span></span> | <span data-ttu-id="dd70d-p119">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="dd70d-404">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-404">String</span></span> | <span data-ttu-id="dd70d-p120">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd70d-407">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-407">Requirements</span></span>

|<span data-ttu-id="dd70d-408">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-408">Requirement</span></span>| <span data-ttu-id="dd70d-409">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-410">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-411">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-411">1.0</span></span>|
|[<span data-ttu-id="dd70d-412">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-413">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-414">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-415">Lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-416">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="dd70d-417">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="dd70d-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="dd70d-418">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="dd70d-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="dd70d-419">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="dd70d-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="dd70d-420">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="dd70d-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="dd70d-421">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="dd70d-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-422">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dd70d-422">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-423">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="dd70d-423">All parameters are optional.</span></span>

|<span data-ttu-id="dd70d-424">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-424">Name</span></span>| <span data-ttu-id="dd70d-425">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-425">Type</span></span>| <span data-ttu-id="dd70d-426">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="dd70d-427">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-427">Object</span></span> | <span data-ttu-id="dd70d-428">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="dd70d-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="dd70d-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="dd70d-430">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="dd70d-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="dd70d-431">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="dd70d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="dd70d-433">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="dd70d-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="dd70d-434">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="dd70d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="dd70d-436">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="dd70d-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="dd70d-437">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="dd70d-438">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-438">String</span></span> | <span data-ttu-id="dd70d-439">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="dd70d-439">A string containing the subject of the message.</span></span> <span data-ttu-id="dd70d-440">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="dd70d-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="dd70d-441">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dd70d-441">String</span></span> | <span data-ttu-id="dd70d-442">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="dd70d-442">The HTML body of the message.</span></span> <span data-ttu-id="dd70d-443">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="dd70d-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="dd70d-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="dd70d-445">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="dd70d-446">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-446">String</span></span> | <span data-ttu-id="dd70d-p127">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="dd70d-449">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-449">String</span></span> | <span data-ttu-id="dd70d-450">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="dd70d-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="dd70d-451">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dd70d-451">String</span></span> | <span data-ttu-id="dd70d-p128">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="dd70d-454">Booléen</span><span class="sxs-lookup"><span data-stu-id="dd70d-454">Boolean</span></span> | <span data-ttu-id="dd70d-p129">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="dd70d-457">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-457">String</span></span> | <span data-ttu-id="dd70d-458">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="dd70d-459">ID d’élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="dd70d-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="dd70d-460">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="dd70d-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="dd70d-461">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-461">Requirements</span></span>

|<span data-ttu-id="dd70d-462">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-462">Requirement</span></span>| <span data-ttu-id="dd70d-463">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-464">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-465">1.6</span><span class="sxs-lookup"><span data-stu-id="dd70d-465">1.6</span></span> |
|[<span data-ttu-id="dd70d-466">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-466">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-467">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-468">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-468">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-469">Lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-470">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-470">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="dd70d-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="dd70d-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="dd70d-472">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="dd70d-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="dd70d-p131">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-475">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="dd70d-475">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="dd70d-476">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dd70d-476">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="dd70d-477">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-477">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="dd70d-478">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dd70d-478">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="dd70d-479">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="dd70d-479">**REST Tokens**</span></span>

<span data-ttu-id="dd70d-p133">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="dd70d-483">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="dd70d-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="dd70d-484">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="dd70d-484">**EWS Tokens**</span></span>

<span data-ttu-id="dd70d-p134">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="dd70d-487">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="dd70d-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="dd70d-488">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="dd70d-488">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="dd70d-489">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-489">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="dd70d-490">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dd70d-490">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-491">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-491">Parameters</span></span>

|<span data-ttu-id="dd70d-492">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-492">Name</span></span>| <span data-ttu-id="dd70d-493">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-493">Type</span></span>| <span data-ttu-id="dd70d-494">Attributs</span><span class="sxs-lookup"><span data-stu-id="dd70d-494">Attributes</span></span>| <span data-ttu-id="dd70d-495">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-495">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="dd70d-496">Object</span><span class="sxs-lookup"><span data-stu-id="dd70d-496">Object</span></span> | <span data-ttu-id="dd70d-497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-497">&lt;optional&gt;</span></span> | <span data-ttu-id="dd70d-498">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="dd70d-498">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="dd70d-499">Boolean</span><span class="sxs-lookup"><span data-stu-id="dd70d-499">Boolean</span></span> |  <span data-ttu-id="dd70d-500">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-500">&lt;optional&gt;</span></span> | <span data-ttu-id="dd70d-p136">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dd70d-503">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-503">Object</span></span> |  <span data-ttu-id="dd70d-504">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-504">&lt;optional&gt;</span></span> | <span data-ttu-id="dd70d-505">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dd70d-505">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="dd70d-506">fonction</span><span class="sxs-lookup"><span data-stu-id="dd70d-506">function</span></span>||<span data-ttu-id="dd70d-507">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd70d-507">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd70d-508">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-508">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dd70d-509">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dd70d-509">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd70d-510">Erreurs</span><span class="sxs-lookup"><span data-stu-id="dd70d-510">Errors</span></span>

|<span data-ttu-id="dd70d-511">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="dd70d-511">Error code</span></span>|<span data-ttu-id="dd70d-512">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-512">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dd70d-513">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="dd70d-513">The request has failed.</span></span> <span data-ttu-id="dd70d-514">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dd70d-514">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dd70d-515">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="dd70d-515">The Exchange server returned an error.</span></span> <span data-ttu-id="dd70d-516">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dd70d-516">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dd70d-517">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="dd70d-517">The user is no longer connected to the network.</span></span> <span data-ttu-id="dd70d-518">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="dd70d-518">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-519">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-519">Requirements</span></span>

|<span data-ttu-id="dd70d-520">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-520">Requirement</span></span>| <span data-ttu-id="dd70d-521">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-522">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-523">1,5</span><span class="sxs-lookup"><span data-stu-id="dd70d-523">1.5</span></span> |
|[<span data-ttu-id="dd70d-524">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-525">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-526">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-527">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-527">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-528">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-528">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="dd70d-529">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dd70d-529">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dd70d-530">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="dd70d-530">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="dd70d-p140">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="dd70d-533">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="dd70d-533">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="dd70d-534">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-534">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="dd70d-535">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="dd70d-535">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="dd70d-536">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="dd70d-536">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="dd70d-537">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-537">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="dd70d-538">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="dd70d-538">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-539">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-539">Parameters</span></span>

|<span data-ttu-id="dd70d-540">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-540">Name</span></span>| <span data-ttu-id="dd70d-541">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-541">Type</span></span>| <span data-ttu-id="dd70d-542">Attributs</span><span class="sxs-lookup"><span data-stu-id="dd70d-542">Attributes</span></span>| <span data-ttu-id="dd70d-543">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-543">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dd70d-544">function</span><span class="sxs-lookup"><span data-stu-id="dd70d-544">function</span></span>||<span data-ttu-id="dd70d-545">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd70d-545">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd70d-546">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-546">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dd70d-547">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dd70d-547">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="dd70d-548">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-548">Object</span></span>| <span data-ttu-id="dd70d-549">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-549">&lt;optional&gt;</span></span>|<span data-ttu-id="dd70d-550">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dd70d-550">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd70d-551">Erreurs</span><span class="sxs-lookup"><span data-stu-id="dd70d-551">Errors</span></span>

|<span data-ttu-id="dd70d-552">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="dd70d-552">Error code</span></span>|<span data-ttu-id="dd70d-553">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-553">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dd70d-554">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="dd70d-554">The request has failed.</span></span> <span data-ttu-id="dd70d-555">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dd70d-555">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dd70d-556">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="dd70d-556">The Exchange server returned an error.</span></span> <span data-ttu-id="dd70d-557">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dd70d-557">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dd70d-558">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="dd70d-558">The user is no longer connected to the network.</span></span> <span data-ttu-id="dd70d-559">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="dd70d-559">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-560">Requirements</span></span>

|<span data-ttu-id="dd70d-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-561">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="dd70d-562">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-563">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-563">1.0</span></span> | <span data-ttu-id="dd70d-564">1.3</span><span class="sxs-lookup"><span data-stu-id="dd70d-564">1.3</span></span> |
|[<span data-ttu-id="dd70d-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-566">ReadItem</span></span> | <span data-ttu-id="dd70d-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-567">ReadItem</span></span> |
|[<span data-ttu-id="dd70d-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-569">Lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-569">Read</span></span> | <span data-ttu-id="dd70d-570">Composition</span><span class="sxs-lookup"><span data-stu-id="dd70d-570">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="dd70d-571">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-571">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="dd70d-572">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dd70d-572">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="dd70d-573">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="dd70d-573">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="dd70d-574">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="dd70d-574">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-575">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dd70d-575">Parameters</span></span>

|<span data-ttu-id="dd70d-576">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-576">Name</span></span>| <span data-ttu-id="dd70d-577">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-577">Type</span></span>| <span data-ttu-id="dd70d-578">Attributs</span><span class="sxs-lookup"><span data-stu-id="dd70d-578">Attributes</span></span>| <span data-ttu-id="dd70d-579">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-579">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="dd70d-580">function</span><span class="sxs-lookup"><span data-stu-id="dd70d-580">function</span></span>||<span data-ttu-id="dd70d-581">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd70d-581">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd70d-582">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-582">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="dd70d-583">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="dd70d-583">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="dd70d-584">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-584">Object</span></span>| <span data-ttu-id="dd70d-585">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-585">&lt;optional&gt;</span></span>|<span data-ttu-id="dd70d-586">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dd70d-586">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="dd70d-587">Erreurs</span><span class="sxs-lookup"><span data-stu-id="dd70d-587">Errors</span></span>

|<span data-ttu-id="dd70d-588">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="dd70d-588">Error code</span></span>|<span data-ttu-id="dd70d-589">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-589">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="dd70d-590">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="dd70d-590">The request has failed.</span></span> <span data-ttu-id="dd70d-591">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dd70d-591">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="dd70d-592">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="dd70d-592">The Exchange server returned an error.</span></span> <span data-ttu-id="dd70d-593">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="dd70d-593">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="dd70d-594">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="dd70d-594">The user is no longer connected to the network.</span></span> <span data-ttu-id="dd70d-595">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="dd70d-595">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-596">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-596">Requirements</span></span>

|<span data-ttu-id="dd70d-597">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-597">Requirement</span></span>| <span data-ttu-id="dd70d-598">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-599">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-600">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-600">1.0</span></span>|
|[<span data-ttu-id="dd70d-601">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-602">ReadItem</span></span>|
|[<span data-ttu-id="dd70d-603">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-604">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-604">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-605">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-605">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="dd70d-606">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="dd70d-606">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="dd70d-607">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dd70d-607">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-608">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="dd70d-608">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="dd70d-609">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="dd70d-609">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="dd70d-610">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="dd70d-610">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="dd70d-611">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dd70d-611">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="dd70d-612">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-612">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="dd70d-613">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="dd70d-613">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="dd70d-614">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-614">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="dd70d-615">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="dd70d-615">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="dd70d-p150">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="dd70d-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="dd70d-618">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="dd70d-618">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="dd70d-619">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="dd70d-619">Version differences</span></span>

<span data-ttu-id="dd70d-620">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-620">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="dd70d-p151">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="dd70d-p151">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-624">Parameters</span><span class="sxs-lookup"><span data-stu-id="dd70d-624">Parameters</span></span>

|<span data-ttu-id="dd70d-625">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-625">Name</span></span>| <span data-ttu-id="dd70d-626">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-626">Type</span></span>| <span data-ttu-id="dd70d-627">Attributs</span><span class="sxs-lookup"><span data-stu-id="dd70d-627">Attributes</span></span>| <span data-ttu-id="dd70d-628">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-628">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="dd70d-629">String</span><span class="sxs-lookup"><span data-stu-id="dd70d-629">String</span></span>||<span data-ttu-id="dd70d-630">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="dd70d-630">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="dd70d-631">function</span><span class="sxs-lookup"><span data-stu-id="dd70d-631">function</span></span>||<span data-ttu-id="dd70d-632">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd70d-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="dd70d-633">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="dd70d-633">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="dd70d-634">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="dd70d-634">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="dd70d-635">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-635">Object</span></span>| <span data-ttu-id="dd70d-636">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-636">&lt;optional&gt;</span></span>|<span data-ttu-id="dd70d-637">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dd70d-637">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-638">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-638">Requirements</span></span>

|<span data-ttu-id="dd70d-639">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-639">Requirement</span></span>| <span data-ttu-id="dd70d-640">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-640">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-641">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-641">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-642">1.0</span><span class="sxs-lookup"><span data-stu-id="dd70d-642">1.0</span></span>|
|[<span data-ttu-id="dd70d-643">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-643">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-644">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="dd70d-644">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="dd70d-645">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-645">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-646">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-646">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dd70d-647">Exemple</span><span class="sxs-lookup"><span data-stu-id="dd70d-647">Example</span></span>

<span data-ttu-id="dd70d-648">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="dd70d-648">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="dd70d-649">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="dd70d-649">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="dd70d-650">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="dd70d-650">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="dd70d-651">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="dd70d-651">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="dd70d-652">Paramètres</span><span class="sxs-lookup"><span data-stu-id="dd70d-652">Parameters</span></span>

| <span data-ttu-id="dd70d-653">Nom</span><span class="sxs-lookup"><span data-stu-id="dd70d-653">Name</span></span> | <span data-ttu-id="dd70d-654">Type</span><span class="sxs-lookup"><span data-stu-id="dd70d-654">Type</span></span> | <span data-ttu-id="dd70d-655">Attributs</span><span class="sxs-lookup"><span data-stu-id="dd70d-655">Attributes</span></span> | <span data-ttu-id="dd70d-656">Description</span><span class="sxs-lookup"><span data-stu-id="dd70d-656">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="dd70d-657">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="dd70d-657">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="dd70d-658">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="dd70d-658">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="dd70d-659">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-659">Object</span></span> | <span data-ttu-id="dd70d-660">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-660">&lt;optional&gt;</span></span> | <span data-ttu-id="dd70d-661">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="dd70d-661">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="dd70d-662">Objet</span><span class="sxs-lookup"><span data-stu-id="dd70d-662">Object</span></span> | <span data-ttu-id="dd70d-663">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-663">&lt;optional&gt;</span></span> | <span data-ttu-id="dd70d-664">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="dd70d-664">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="dd70d-665">fonction</span><span class="sxs-lookup"><span data-stu-id="dd70d-665">function</span></span>| <span data-ttu-id="dd70d-666">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="dd70d-666">&lt;optional&gt;</span></span>|<span data-ttu-id="dd70d-667">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="dd70d-667">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd70d-668">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd70d-668">Requirements</span></span>

|<span data-ttu-id="dd70d-669">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd70d-669">Requirement</span></span>| <span data-ttu-id="dd70d-670">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd70d-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd70d-671">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd70d-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="dd70d-672">1,5</span><span class="sxs-lookup"><span data-stu-id="dd70d-672">1.5</span></span> |
|[<span data-ttu-id="dd70d-673">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dd70d-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="dd70d-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="dd70d-674">ReadItem</span></span> |
|[<span data-ttu-id="dd70d-675">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd70d-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="dd70d-676">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd70d-676">Compose or Read</span></span>|
