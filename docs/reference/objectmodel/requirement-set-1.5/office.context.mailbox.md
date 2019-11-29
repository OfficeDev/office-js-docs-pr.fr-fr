---
title: Office.context – ensemble de conditions requises 1.5
description: ''
ms.date: 11/27/2019
localization_priority: Priority
ms.openlocfilehash: eefeab2cf6fbe78451afae7e588640fe7f50dba4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629685"
---
# <a name="mailbox"></a><span data-ttu-id="be59e-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="be59e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="be59e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="be59e-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="be59e-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="be59e-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-105">Requirements</span></span>

|<span data-ttu-id="be59e-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-106">Requirement</span></span>| <span data-ttu-id="be59e-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-109">1.0</span></span>|
|[<span data-ttu-id="be59e-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="be59e-111">Restricted</span></span>|
|[<span data-ttu-id="be59e-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="be59e-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="be59e-114">Members and methods</span></span>

| <span data-ttu-id="be59e-115">Membre</span><span class="sxs-lookup"><span data-stu-id="be59e-115">Member</span></span> | <span data-ttu-id="be59e-116">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="be59e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="be59e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="be59e-118">Membre</span><span class="sxs-lookup"><span data-stu-id="be59e-118">Member</span></span> |
| [<span data-ttu-id="be59e-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="be59e-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="be59e-120">Membre</span><span class="sxs-lookup"><span data-stu-id="be59e-120">Member</span></span> |
| [<span data-ttu-id="be59e-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="be59e-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="be59e-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-122">Method</span></span> |
| [<span data-ttu-id="be59e-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="be59e-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="be59e-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-124">Method</span></span> |
| [<span data-ttu-id="be59e-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="be59e-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="be59e-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-126">Method</span></span> |
| [<span data-ttu-id="be59e-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="be59e-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="be59e-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-128">Method</span></span> |
| [<span data-ttu-id="be59e-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="be59e-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="be59e-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-130">Method</span></span> |
| [<span data-ttu-id="be59e-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="be59e-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="be59e-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-132">Method</span></span> |
| [<span data-ttu-id="be59e-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="be59e-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="be59e-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-134">Method</span></span> |
| [<span data-ttu-id="be59e-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="be59e-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="be59e-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-136">Method</span></span> |
| [<span data-ttu-id="be59e-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="be59e-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="be59e-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-138">Method</span></span> |
| [<span data-ttu-id="be59e-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="be59e-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="be59e-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-140">Method</span></span> |
| [<span data-ttu-id="be59e-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="be59e-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="be59e-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-142">Method</span></span> |
| [<span data-ttu-id="be59e-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="be59e-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="be59e-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-144">Method</span></span> |
| [<span data-ttu-id="be59e-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="be59e-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="be59e-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="be59e-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="be59e-147">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="be59e-147">Namespaces</span></span>

<span data-ttu-id="be59e-148">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="be59e-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="be59e-149">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="be59e-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="be59e-150">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="be59e-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="be59e-151">Members</span><span class="sxs-lookup"><span data-stu-id="be59e-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="be59e-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="be59e-152">ewsUrl: String</span></span>

<span data-ttu-id="be59e-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="be59e-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-155">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="be59e-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="be59e-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="be59e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="be59e-158">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="be59e-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="be59e-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="be59e-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="be59e-161">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-161">Type</span></span>

*   <span data-ttu-id="be59e-162">String</span><span class="sxs-lookup"><span data-stu-id="be59e-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be59e-163">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-163">Requirements</span></span>

|<span data-ttu-id="be59e-164">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-164">Requirement</span></span>| <span data-ttu-id="be59e-165">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-166">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-167">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-167">1.0</span></span>|
|[<span data-ttu-id="be59e-168">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-169">ReadItem</span></span>|
|[<span data-ttu-id="be59e-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="be59e-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="be59e-172">restUrl: String</span></span>

<span data-ttu-id="be59e-173">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="be59e-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="be59e-174">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be59e-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-175">Les clients Outlook connectés aux installations locales d’Exchange 2016 ou version ultérieure avec une URL REST personnalisée configurée renvoient une valeur non valide pour `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="be59e-175">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="be59e-176">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-176">Type</span></span>

*   <span data-ttu-id="be59e-177">String</span><span class="sxs-lookup"><span data-stu-id="be59e-177">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="be59e-178">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-178">Requirements</span></span>

|<span data-ttu-id="be59e-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-179">Requirement</span></span>| <span data-ttu-id="be59e-180">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-181">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-182">1,5</span><span class="sxs-lookup"><span data-stu-id="be59e-182">1.5</span></span> |
|[<span data-ttu-id="be59e-183">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-184">ReadItem</span></span>|
|[<span data-ttu-id="be59e-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-186">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-186">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="be59e-187">Méthodes</span><span class="sxs-lookup"><span data-stu-id="be59e-187">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="be59e-188">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="be59e-188">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="be59e-189">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="be59e-189">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="be59e-190">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé quand l’utilisateur sélectionne un nouvel élément.</span><span class="sxs-lookup"><span data-stu-id="be59e-190">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="be59e-191">Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’IU du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="be59e-191">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-192">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-192">Parameters</span></span>

| <span data-ttu-id="be59e-193">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-193">Name</span></span> | <span data-ttu-id="be59e-194">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-194">Type</span></span> | <span data-ttu-id="be59e-195">Attributs</span><span class="sxs-lookup"><span data-stu-id="be59e-195">Attributes</span></span> | <span data-ttu-id="be59e-196">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="be59e-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="be59e-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="be59e-198">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="be59e-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="be59e-199">Fonction</span><span class="sxs-lookup"><span data-stu-id="be59e-199">Function</span></span> || <span data-ttu-id="be59e-p105">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="be59e-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="be59e-203">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-203">Object</span></span> | <span data-ttu-id="be59e-204">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-204">&lt;optional&gt;</span></span> | <span data-ttu-id="be59e-205">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="be59e-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="be59e-206">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-206">Object</span></span> | <span data-ttu-id="be59e-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-207">&lt;optional&gt;</span></span> | <span data-ttu-id="be59e-208">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="be59e-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="be59e-209">fonction</span><span class="sxs-lookup"><span data-stu-id="be59e-209">function</span></span>| <span data-ttu-id="be59e-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-210">&lt;optional&gt;</span></span>|<span data-ttu-id="be59e-211">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="be59e-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-212">Requirements</span></span>

|<span data-ttu-id="be59e-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-213">Requirement</span></span>| <span data-ttu-id="be59e-214">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-216">1,5</span><span class="sxs-lookup"><span data-stu-id="be59e-216">1.5</span></span> |
|[<span data-ttu-id="be59e-217">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-218">ReadItem</span></span> |
|[<span data-ttu-id="be59e-219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-220">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be59e-221">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-221">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="be59e-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="be59e-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="be59e-223">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="be59e-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-224">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="be59e-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="be59e-p106">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="be59e-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-227">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-227">Parameters</span></span>

|<span data-ttu-id="be59e-228">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-228">Name</span></span>| <span data-ttu-id="be59e-229">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-229">Type</span></span>| <span data-ttu-id="be59e-230">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="be59e-231">String</span><span class="sxs-lookup"><span data-stu-id="be59e-231">String</span></span>|<span data-ttu-id="be59e-232">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="be59e-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="be59e-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="be59e-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="be59e-234">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="be59e-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-235">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-235">Requirements</span></span>

|<span data-ttu-id="be59e-236">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-236">Requirement</span></span>| <span data-ttu-id="be59e-237">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-238">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-239">1.3</span><span class="sxs-lookup"><span data-stu-id="be59e-239">1.3</span></span>|
|[<span data-ttu-id="be59e-240">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-241">Restreinte</span><span class="sxs-lookup"><span data-stu-id="be59e-241">Restricted</span></span>|
|[<span data-ttu-id="be59e-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be59e-244">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="be59e-244">Returns:</span></span>

<span data-ttu-id="be59e-245">Type : String</span><span class="sxs-lookup"><span data-stu-id="be59e-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="be59e-246">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="be59e-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="be59e-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="be59e-248">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="be59e-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="be59e-p107">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be59e-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="be59e-p108">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="be59e-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-254">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-254">Parameters</span></span>

|<span data-ttu-id="be59e-255">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-255">Name</span></span>| <span data-ttu-id="be59e-256">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-256">Type</span></span>| <span data-ttu-id="be59e-257">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="be59e-258">Date</span><span class="sxs-lookup"><span data-stu-id="be59e-258">Date</span></span>|<span data-ttu-id="be59e-259">Objet Date</span><span class="sxs-lookup"><span data-stu-id="be59e-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-260">Requirements</span></span>

|<span data-ttu-id="be59e-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-261">Requirement</span></span>| <span data-ttu-id="be59e-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-264">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-264">1.0</span></span>|
|[<span data-ttu-id="be59e-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-266">ReadItem</span></span>|
|[<span data-ttu-id="be59e-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be59e-269">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="be59e-269">Returns:</span></span>

<span data-ttu-id="be59e-270">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="be59e-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="be59e-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="be59e-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="be59e-272">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="be59e-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-273">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="be59e-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="be59e-p109">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="be59e-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-276">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-276">Parameters</span></span>

|<span data-ttu-id="be59e-277">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-277">Name</span></span>| <span data-ttu-id="be59e-278">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-278">Type</span></span>| <span data-ttu-id="be59e-279">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="be59e-280">String</span><span class="sxs-lookup"><span data-stu-id="be59e-280">String</span></span>|<span data-ttu-id="be59e-281">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="be59e-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="be59e-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="be59e-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="be59e-283">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="be59e-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-284">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-284">Requirements</span></span>

|<span data-ttu-id="be59e-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-285">Requirement</span></span>| <span data-ttu-id="be59e-286">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-287">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-288">1.3</span><span class="sxs-lookup"><span data-stu-id="be59e-288">1.3</span></span>|
|[<span data-ttu-id="be59e-289">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-290">Restreinte</span><span class="sxs-lookup"><span data-stu-id="be59e-290">Restricted</span></span>|
|[<span data-ttu-id="be59e-291">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-292">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be59e-293">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="be59e-293">Returns:</span></span>

<span data-ttu-id="be59e-294">Type : String</span><span class="sxs-lookup"><span data-stu-id="be59e-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="be59e-295">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="be59e-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="be59e-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="be59e-297">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="be59e-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="be59e-298">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="be59e-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-299">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-299">Parameters</span></span>

|<span data-ttu-id="be59e-300">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-300">Name</span></span>| <span data-ttu-id="be59e-301">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-301">Type</span></span>| <span data-ttu-id="be59e-302">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="be59e-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="be59e-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="be59e-304">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="be59e-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-305">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-305">Requirements</span></span>

|<span data-ttu-id="be59e-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-306">Requirement</span></span>| <span data-ttu-id="be59e-307">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-308">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-309">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-309">1.0</span></span>|
|[<span data-ttu-id="be59e-310">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-311">ReadItem</span></span>|
|[<span data-ttu-id="be59e-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-313">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="be59e-314">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="be59e-314">Returns:</span></span>

<span data-ttu-id="be59e-315">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="be59e-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="be59e-316">Type : Date</span><span class="sxs-lookup"><span data-stu-id="be59e-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="be59e-317">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="be59e-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="be59e-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="be59e-319">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="be59e-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-320">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="be59e-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="be59e-321">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="be59e-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="be59e-p110">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="be59e-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="be59e-324">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="be59e-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="be59e-325">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="be59e-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-326">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-326">Parameters</span></span>

|<span data-ttu-id="be59e-327">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-327">Name</span></span>| <span data-ttu-id="be59e-328">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-328">Type</span></span>| <span data-ttu-id="be59e-329">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="be59e-330">String</span><span class="sxs-lookup"><span data-stu-id="be59e-330">String</span></span>|<span data-ttu-id="be59e-331">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="be59e-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-332">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-332">Requirements</span></span>

|<span data-ttu-id="be59e-333">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-333">Requirement</span></span>| <span data-ttu-id="be59e-334">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-335">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-336">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-336">1.0</span></span>|
|[<span data-ttu-id="be59e-337">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-338">ReadItem</span></span>|
|[<span data-ttu-id="be59e-339">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-340">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be59e-341">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="be59e-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="be59e-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="be59e-343">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="be59e-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-344">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="be59e-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="be59e-345">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="be59e-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="be59e-346">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="be59e-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="be59e-347">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="be59e-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="be59e-p111">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="be59e-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-350">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-350">Parameters</span></span>

|<span data-ttu-id="be59e-351">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-351">Name</span></span>| <span data-ttu-id="be59e-352">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-352">Type</span></span>| <span data-ttu-id="be59e-353">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="be59e-354">String</span><span class="sxs-lookup"><span data-stu-id="be59e-354">String</span></span>|<span data-ttu-id="be59e-355">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="be59e-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-356">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-356">Requirements</span></span>

|<span data-ttu-id="be59e-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-357">Requirement</span></span>| <span data-ttu-id="be59e-358">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-359">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-360">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-360">1.0</span></span>|
|[<span data-ttu-id="be59e-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-362">ReadItem</span></span>|
|[<span data-ttu-id="be59e-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-364">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be59e-365">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="be59e-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="be59e-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="be59e-367">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="be59e-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-368">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="be59e-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="be59e-p112">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="be59e-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="be59e-p113">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="be59e-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="be59e-p114">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="be59e-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="be59e-376">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="be59e-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-377">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-377">Parameters</span></span>

|<span data-ttu-id="be59e-378">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-378">Name</span></span>| <span data-ttu-id="be59e-379">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-379">Type</span></span>| <span data-ttu-id="be59e-380">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-380">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="be59e-381">Object</span><span class="sxs-lookup"><span data-stu-id="be59e-381">Object</span></span> | <span data-ttu-id="be59e-382">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="be59e-382">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="be59e-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="be59e-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="be59e-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="be59e-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="be59e-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="be59e-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="be59e-389">Date</span><span class="sxs-lookup"><span data-stu-id="be59e-389">Date</span></span> | <span data-ttu-id="be59e-390">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="be59e-390">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="be59e-391">Date</span><span class="sxs-lookup"><span data-stu-id="be59e-391">Date</span></span> | <span data-ttu-id="be59e-392">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="be59e-392">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="be59e-393">String</span><span class="sxs-lookup"><span data-stu-id="be59e-393">String</span></span> | <span data-ttu-id="be59e-p117">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="be59e-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="be59e-396">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-396">Array.&lt;String&gt;</span></span> | <span data-ttu-id="be59e-p118">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="be59e-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="be59e-399">String</span><span class="sxs-lookup"><span data-stu-id="be59e-399">String</span></span> | <span data-ttu-id="be59e-p119">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="be59e-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="be59e-402">String</span><span class="sxs-lookup"><span data-stu-id="be59e-402">String</span></span> | <span data-ttu-id="be59e-p120">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="be59e-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="be59e-405">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-405">Requirements</span></span>

|<span data-ttu-id="be59e-406">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-406">Requirement</span></span>| <span data-ttu-id="be59e-407">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-408">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-409">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-409">1.0</span></span>|
|[<span data-ttu-id="be59e-410">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-411">ReadItem</span></span>|
|[<span data-ttu-id="be59e-412">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-413">Lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be59e-414">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-414">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="be59e-415">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="be59e-415">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="be59e-416">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="be59e-416">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="be59e-p121">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="be59e-p121">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-419">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="be59e-419">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="be59e-420">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="be59e-420">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="be59e-421">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="be59e-421">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="be59e-422">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="be59e-422">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="be59e-423">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="be59e-423">**REST Tokens**</span></span>

<span data-ttu-id="be59e-p123">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="be59e-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="be59e-427">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="be59e-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="be59e-428">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="be59e-428">**EWS Tokens**</span></span>

<span data-ttu-id="be59e-p124">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="be59e-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="be59e-431">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="be59e-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="be59e-432">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="be59e-432">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="be59e-433">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="be59e-433">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="be59e-434">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="be59e-434">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-435">Parameters</span><span class="sxs-lookup"><span data-stu-id="be59e-435">Parameters</span></span>

|<span data-ttu-id="be59e-436">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-436">Name</span></span>| <span data-ttu-id="be59e-437">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-437">Type</span></span>| <span data-ttu-id="be59e-438">Attributs</span><span class="sxs-lookup"><span data-stu-id="be59e-438">Attributes</span></span>| <span data-ttu-id="be59e-439">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-439">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="be59e-440">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-440">Object</span></span> | <span data-ttu-id="be59e-441">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-441">&lt;optional&gt;</span></span> | <span data-ttu-id="be59e-442">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="be59e-442">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="be59e-443">Boolean</span><span class="sxs-lookup"><span data-stu-id="be59e-443">Boolean</span></span> |  <span data-ttu-id="be59e-444">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-444">&lt;optional&gt;</span></span> | <span data-ttu-id="be59e-p126">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="be59e-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="be59e-447">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-447">Object</span></span> |  <span data-ttu-id="be59e-448">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-448">&lt;optional&gt;</span></span> | <span data-ttu-id="be59e-449">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="be59e-449">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="be59e-450">fonction</span><span class="sxs-lookup"><span data-stu-id="be59e-450">function</span></span>||<span data-ttu-id="be59e-451">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="be59e-451">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="be59e-452">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="be59e-452">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="be59e-453">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="be59e-453">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="be59e-454">Erreurs</span><span class="sxs-lookup"><span data-stu-id="be59e-454">Errors</span></span>

|<span data-ttu-id="be59e-455">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="be59e-455">Error code</span></span>|<span data-ttu-id="be59e-456">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-456">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="be59e-457">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="be59e-457">The request has failed.</span></span> <span data-ttu-id="be59e-458">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="be59e-458">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="be59e-459">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="be59e-459">The Exchange server returned an error.</span></span> <span data-ttu-id="be59e-460">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="be59e-460">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="be59e-461">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="be59e-461">The user is no longer connected to the network.</span></span> <span data-ttu-id="be59e-462">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="be59e-462">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-463">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-463">Requirements</span></span>

|<span data-ttu-id="be59e-464">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-464">Requirement</span></span>| <span data-ttu-id="be59e-465">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-466">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-467">1,5</span><span class="sxs-lookup"><span data-stu-id="be59e-467">1.5</span></span> |
|[<span data-ttu-id="be59e-468">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-469">ReadItem</span></span>|
|[<span data-ttu-id="be59e-470">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-471">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-471">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="be59e-472">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-472">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="be59e-473">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="be59e-473">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="be59e-474">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="be59e-474">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="be59e-p130">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="be59e-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="be59e-477">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="be59e-477">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="be59e-478">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="be59e-478">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="be59e-479">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="be59e-479">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="be59e-480">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="be59e-480">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="be59e-481">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="be59e-481">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="be59e-482">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="be59e-482">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-483">Parameters</span><span class="sxs-lookup"><span data-stu-id="be59e-483">Parameters</span></span>

|<span data-ttu-id="be59e-484">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-484">Name</span></span>| <span data-ttu-id="be59e-485">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-485">Type</span></span>| <span data-ttu-id="be59e-486">Attributs</span><span class="sxs-lookup"><span data-stu-id="be59e-486">Attributes</span></span>| <span data-ttu-id="be59e-487">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-487">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="be59e-488">function</span><span class="sxs-lookup"><span data-stu-id="be59e-488">function</span></span>||<span data-ttu-id="be59e-489">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="be59e-489">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="be59e-490">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="be59e-490">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="be59e-491">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="be59e-491">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="be59e-492">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-492">Object</span></span>| <span data-ttu-id="be59e-493">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-493">&lt;optional&gt;</span></span>|<span data-ttu-id="be59e-494">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="be59e-494">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="be59e-495">Erreurs</span><span class="sxs-lookup"><span data-stu-id="be59e-495">Errors</span></span>

|<span data-ttu-id="be59e-496">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="be59e-496">Error code</span></span>|<span data-ttu-id="be59e-497">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-497">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="be59e-498">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="be59e-498">The request has failed.</span></span> <span data-ttu-id="be59e-499">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="be59e-499">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="be59e-500">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="be59e-500">The Exchange server returned an error.</span></span> <span data-ttu-id="be59e-501">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="be59e-501">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="be59e-502">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="be59e-502">The user is no longer connected to the network.</span></span> <span data-ttu-id="be59e-503">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="be59e-503">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-504">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-504">Requirements</span></span>

|<span data-ttu-id="be59e-505">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-505">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="be59e-506">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-507">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-507">1.0</span></span> | <span data-ttu-id="be59e-508">1.3</span><span class="sxs-lookup"><span data-stu-id="be59e-508">1.3</span></span> |
|[<span data-ttu-id="be59e-509">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-510">ReadItem</span></span> | <span data-ttu-id="be59e-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-511">ReadItem</span></span> |
|[<span data-ttu-id="be59e-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-513">Lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-513">Read</span></span> | <span data-ttu-id="be59e-514">Composition</span><span class="sxs-lookup"><span data-stu-id="be59e-514">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="be59e-515">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-515">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="be59e-516">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="be59e-516">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="be59e-517">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="be59e-517">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="be59e-518">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="be59e-518">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-519">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-519">Parameters</span></span>

|<span data-ttu-id="be59e-520">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-520">Name</span></span>| <span data-ttu-id="be59e-521">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-521">Type</span></span>| <span data-ttu-id="be59e-522">Attributs</span><span class="sxs-lookup"><span data-stu-id="be59e-522">Attributes</span></span>| <span data-ttu-id="be59e-523">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-523">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="be59e-524">function</span><span class="sxs-lookup"><span data-stu-id="be59e-524">function</span></span>||<span data-ttu-id="be59e-525">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="be59e-525">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="be59e-526">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="be59e-526">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="be59e-527">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="be59e-527">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="be59e-528">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-528">Object</span></span>| <span data-ttu-id="be59e-529">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-529">&lt;optional&gt;</span></span>|<span data-ttu-id="be59e-530">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="be59e-530">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="be59e-531">Erreurs</span><span class="sxs-lookup"><span data-stu-id="be59e-531">Errors</span></span>

|<span data-ttu-id="be59e-532">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="be59e-532">Error code</span></span>|<span data-ttu-id="be59e-533">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-533">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="be59e-534">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="be59e-534">The request has failed.</span></span> <span data-ttu-id="be59e-535">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="be59e-535">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="be59e-536">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="be59e-536">The Exchange server returned an error.</span></span> <span data-ttu-id="be59e-537">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="be59e-537">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="be59e-538">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="be59e-538">The user is no longer connected to the network.</span></span> <span data-ttu-id="be59e-539">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="be59e-539">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-540">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-540">Requirements</span></span>

|<span data-ttu-id="be59e-541">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-541">Requirement</span></span>| <span data-ttu-id="be59e-542">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-543">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-544">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-544">1.0</span></span>|
|[<span data-ttu-id="be59e-545">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-545">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-546">ReadItem</span></span>|
|[<span data-ttu-id="be59e-547">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-547">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-548">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-548">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be59e-549">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-549">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="be59e-550">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="be59e-550">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="be59e-551">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be59e-551">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-552">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="be59e-552">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="be59e-553">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="be59e-553">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="be59e-554">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="be59e-554">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="be59e-555">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="be59e-555">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="be59e-556">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="be59e-556">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="be59e-557">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="be59e-557">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="be59e-558">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="be59e-558">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="be59e-559">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="be59e-559">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="be59e-p140">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="be59e-p140">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="be59e-562">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="be59e-562">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="be59e-563">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="be59e-563">Version differences</span></span>

<span data-ttu-id="be59e-564">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="be59e-564">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="be59e-p141">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="be59e-p141">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-568">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-568">Parameters</span></span>

|<span data-ttu-id="be59e-569">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-569">Name</span></span>| <span data-ttu-id="be59e-570">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-570">Type</span></span>| <span data-ttu-id="be59e-571">Attributs</span><span class="sxs-lookup"><span data-stu-id="be59e-571">Attributes</span></span>| <span data-ttu-id="be59e-572">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-572">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="be59e-573">String</span><span class="sxs-lookup"><span data-stu-id="be59e-573">String</span></span>||<span data-ttu-id="be59e-574">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="be59e-574">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="be59e-575">function</span><span class="sxs-lookup"><span data-stu-id="be59e-575">function</span></span>||<span data-ttu-id="be59e-576">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="be59e-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="be59e-577">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="be59e-577">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="be59e-578">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="be59e-578">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="be59e-579">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-579">Object</span></span>| <span data-ttu-id="be59e-580">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-580">&lt;optional&gt;</span></span>|<span data-ttu-id="be59e-581">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="be59e-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-582">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-582">Requirements</span></span>

|<span data-ttu-id="be59e-583">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-583">Requirement</span></span>| <span data-ttu-id="be59e-584">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-585">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-586">1.0</span><span class="sxs-lookup"><span data-stu-id="be59e-586">1.0</span></span>|
|[<span data-ttu-id="be59e-587">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-588">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="be59e-588">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="be59e-589">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-590">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-590">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="be59e-591">Exemple</span><span class="sxs-lookup"><span data-stu-id="be59e-591">Example</span></span>

<span data-ttu-id="be59e-592">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="be59e-592">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="be59e-593">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="be59e-593">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="be59e-594">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="be59e-594">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="be59e-595">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="be59e-595">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="be59e-596">Paramètres</span><span class="sxs-lookup"><span data-stu-id="be59e-596">Parameters</span></span>

| <span data-ttu-id="be59e-597">Nom</span><span class="sxs-lookup"><span data-stu-id="be59e-597">Name</span></span> | <span data-ttu-id="be59e-598">Type</span><span class="sxs-lookup"><span data-stu-id="be59e-598">Type</span></span> | <span data-ttu-id="be59e-599">Attributs</span><span class="sxs-lookup"><span data-stu-id="be59e-599">Attributes</span></span> | <span data-ttu-id="be59e-600">Description</span><span class="sxs-lookup"><span data-stu-id="be59e-600">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="be59e-601">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="be59e-601">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="be59e-602">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="be59e-602">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="be59e-603">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-603">Object</span></span> | <span data-ttu-id="be59e-604">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-604">&lt;optional&gt;</span></span> | <span data-ttu-id="be59e-605">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="be59e-605">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="be59e-606">Objet</span><span class="sxs-lookup"><span data-stu-id="be59e-606">Object</span></span> | <span data-ttu-id="be59e-607">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-607">&lt;optional&gt;</span></span> | <span data-ttu-id="be59e-608">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="be59e-608">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="be59e-609">fonction</span><span class="sxs-lookup"><span data-stu-id="be59e-609">function</span></span>| <span data-ttu-id="be59e-610">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="be59e-610">&lt;optional&gt;</span></span>|<span data-ttu-id="be59e-611">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="be59e-611">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="be59e-612">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="be59e-612">Requirements</span></span>

|<span data-ttu-id="be59e-613">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="be59e-613">Requirement</span></span>| <span data-ttu-id="be59e-614">Valeur</span><span class="sxs-lookup"><span data-stu-id="be59e-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="be59e-615">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="be59e-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="be59e-616">1,5</span><span class="sxs-lookup"><span data-stu-id="be59e-616">1.5</span></span> |
|[<span data-ttu-id="be59e-617">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="be59e-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="be59e-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="be59e-618">ReadItem</span></span> |
|[<span data-ttu-id="be59e-619">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="be59e-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="be59e-620">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="be59e-620">Compose or Read</span></span>|
