---
title: Office.context – ensemble de conditions requises 1.5
description: ''
ms.date: 10/21/2019
localization_priority: Priority
ms.openlocfilehash: bb63d8186d41d072aa62b180b16958d61ce9a66c
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627012"
---
# <a name="mailbox"></a><span data-ttu-id="894db-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="894db-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="894db-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="894db-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="894db-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="894db-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-105">Requirements</span></span>

|<span data-ttu-id="894db-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-106">Requirement</span></span>| <span data-ttu-id="894db-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-109">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-109">1.0</span></span>|
|[<span data-ttu-id="894db-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="894db-111">Restricted</span></span>|
|[<span data-ttu-id="894db-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="894db-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="894db-114">Members and methods</span></span>

| <span data-ttu-id="894db-115">Membre</span><span class="sxs-lookup"><span data-stu-id="894db-115">Member</span></span> | <span data-ttu-id="894db-116">Type</span><span class="sxs-lookup"><span data-stu-id="894db-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="894db-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="894db-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="894db-118">Membre</span><span class="sxs-lookup"><span data-stu-id="894db-118">Member</span></span> |
| [<span data-ttu-id="894db-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="894db-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="894db-120">Membre</span><span class="sxs-lookup"><span data-stu-id="894db-120">Member</span></span> |
| [<span data-ttu-id="894db-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="894db-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="894db-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-122">Method</span></span> |
| [<span data-ttu-id="894db-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="894db-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="894db-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-124">Method</span></span> |
| [<span data-ttu-id="894db-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="894db-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="894db-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-126">Method</span></span> |
| [<span data-ttu-id="894db-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="894db-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="894db-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-128">Method</span></span> |
| [<span data-ttu-id="894db-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="894db-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="894db-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-130">Method</span></span> |
| [<span data-ttu-id="894db-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="894db-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="894db-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-132">Method</span></span> |
| [<span data-ttu-id="894db-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="894db-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="894db-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-134">Method</span></span> |
| [<span data-ttu-id="894db-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="894db-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="894db-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-136">Method</span></span> |
| [<span data-ttu-id="894db-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="894db-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="894db-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-138">Method</span></span> |
| [<span data-ttu-id="894db-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="894db-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="894db-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-140">Method</span></span> |
| [<span data-ttu-id="894db-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="894db-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="894db-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-142">Method</span></span> |
| [<span data-ttu-id="894db-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="894db-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="894db-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-144">Method</span></span> |
| [<span data-ttu-id="894db-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="894db-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="894db-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="894db-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="894db-147">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="894db-147">Namespaces</span></span>

<span data-ttu-id="894db-148">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="894db-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="894db-149">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="894db-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="894db-150">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="894db-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="894db-151">Members</span><span class="sxs-lookup"><span data-stu-id="894db-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="894db-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="894db-152">ewsUrl: String</span></span>

<span data-ttu-id="894db-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="894db-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-155">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="894db-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="894db-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="894db-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="894db-158">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="894db-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="894db-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="894db-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="894db-161">Type</span><span class="sxs-lookup"><span data-stu-id="894db-161">Type</span></span>

*   <span data-ttu-id="894db-162">String</span><span class="sxs-lookup"><span data-stu-id="894db-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="894db-163">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-163">Requirements</span></span>

|<span data-ttu-id="894db-164">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-164">Requirement</span></span>| <span data-ttu-id="894db-165">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-166">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-167">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-167">1.0</span></span>|
|[<span data-ttu-id="894db-168">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-169">ReadItem</span></span>|
|[<span data-ttu-id="894db-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="894db-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="894db-172">restUrl: String</span></span>

<span data-ttu-id="894db-173">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="894db-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="894db-174">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="894db-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="894db-175">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="894db-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="894db-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="894db-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-178">Les clients Outlook connectés aux installations locales d’Exchange 2016 ou version ultérieure avec une URL REST personnalisée configurée renvoient une valeur non valide pour `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="894db-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="894db-179">Type</span><span class="sxs-lookup"><span data-stu-id="894db-179">Type</span></span>

*   <span data-ttu-id="894db-180">String</span><span class="sxs-lookup"><span data-stu-id="894db-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="894db-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-181">Requirements</span></span>

|<span data-ttu-id="894db-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-182">Requirement</span></span>| <span data-ttu-id="894db-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-185">1,5</span><span class="sxs-lookup"><span data-stu-id="894db-185">1.5</span></span> |
|[<span data-ttu-id="894db-186">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-187">ReadItem</span></span>|
|[<span data-ttu-id="894db-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="894db-190">Méthodes</span><span class="sxs-lookup"><span data-stu-id="894db-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="894db-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="894db-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="894db-192">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="894db-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="894db-193">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé quand l’utilisateur sélectionne un nouvel élément.</span><span class="sxs-lookup"><span data-stu-id="894db-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="894db-194">Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’IU du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="894db-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-195">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-195">Parameters</span></span>

| <span data-ttu-id="894db-196">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-196">Name</span></span> | <span data-ttu-id="894db-197">Type</span><span class="sxs-lookup"><span data-stu-id="894db-197">Type</span></span> | <span data-ttu-id="894db-198">Attributs</span><span class="sxs-lookup"><span data-stu-id="894db-198">Attributes</span></span> | <span data-ttu-id="894db-199">Description</span><span class="sxs-lookup"><span data-stu-id="894db-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="894db-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="894db-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="894db-201">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="894db-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="894db-202">Fonction</span><span class="sxs-lookup"><span data-stu-id="894db-202">Function</span></span> || <span data-ttu-id="894db-p106">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="894db-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="894db-206">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-206">Object</span></span> | <span data-ttu-id="894db-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-207">&lt;optional&gt;</span></span> | <span data-ttu-id="894db-208">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="894db-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="894db-209">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-209">Object</span></span> | <span data-ttu-id="894db-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-210">&lt;optional&gt;</span></span> | <span data-ttu-id="894db-211">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="894db-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="894db-212">fonction</span><span class="sxs-lookup"><span data-stu-id="894db-212">function</span></span>| <span data-ttu-id="894db-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-213">&lt;optional&gt;</span></span>|<span data-ttu-id="894db-214">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="894db-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-215">Requirements</span></span>

|<span data-ttu-id="894db-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-216">Requirement</span></span>| <span data-ttu-id="894db-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-219">1,5</span><span class="sxs-lookup"><span data-stu-id="894db-219">1.5</span></span> |
|[<span data-ttu-id="894db-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-221">ReadItem</span></span> |
|[<span data-ttu-id="894db-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-223">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="894db-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="894db-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="894db-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="894db-226">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="894db-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-227">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="894db-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="894db-p107">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="894db-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-230">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-230">Parameters</span></span>

|<span data-ttu-id="894db-231">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-231">Name</span></span>| <span data-ttu-id="894db-232">Type</span><span class="sxs-lookup"><span data-stu-id="894db-232">Type</span></span>| <span data-ttu-id="894db-233">Description</span><span class="sxs-lookup"><span data-stu-id="894db-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="894db-234">String</span><span class="sxs-lookup"><span data-stu-id="894db-234">String</span></span>|<span data-ttu-id="894db-235">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="894db-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="894db-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="894db-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="894db-237">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="894db-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-238">Requirements</span></span>

|<span data-ttu-id="894db-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-239">Requirement</span></span>| <span data-ttu-id="894db-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-242">1.3</span><span class="sxs-lookup"><span data-stu-id="894db-242">1.3</span></span>|
|[<span data-ttu-id="894db-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-244">Restreinte</span><span class="sxs-lookup"><span data-stu-id="894db-244">Restricted</span></span>|
|[<span data-ttu-id="894db-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="894db-247">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="894db-247">Returns:</span></span>

<span data-ttu-id="894db-248">Type : String</span><span class="sxs-lookup"><span data-stu-id="894db-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="894db-249">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="894db-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="894db-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="894db-251">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="894db-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="894db-p108">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="894db-p108">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="894db-p109">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="894db-p109">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-257">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-257">Parameters</span></span>

|<span data-ttu-id="894db-258">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-258">Name</span></span>| <span data-ttu-id="894db-259">Type</span><span class="sxs-lookup"><span data-stu-id="894db-259">Type</span></span>| <span data-ttu-id="894db-260">Description</span><span class="sxs-lookup"><span data-stu-id="894db-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="894db-261">Date</span><span class="sxs-lookup"><span data-stu-id="894db-261">Date</span></span>|<span data-ttu-id="894db-262">Objet Date</span><span class="sxs-lookup"><span data-stu-id="894db-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-263">Requirements</span></span>

|<span data-ttu-id="894db-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-264">Requirement</span></span>| <span data-ttu-id="894db-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-267">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-267">1.0</span></span>|
|[<span data-ttu-id="894db-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-269">ReadItem</span></span>|
|[<span data-ttu-id="894db-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-271">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="894db-272">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="894db-272">Returns:</span></span>

<span data-ttu-id="894db-273">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="894db-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="894db-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="894db-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="894db-275">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="894db-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-276">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="894db-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="894db-p110">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="894db-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-279">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-279">Parameters</span></span>

|<span data-ttu-id="894db-280">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-280">Name</span></span>| <span data-ttu-id="894db-281">Type</span><span class="sxs-lookup"><span data-stu-id="894db-281">Type</span></span>| <span data-ttu-id="894db-282">Description</span><span class="sxs-lookup"><span data-stu-id="894db-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="894db-283">String</span><span class="sxs-lookup"><span data-stu-id="894db-283">String</span></span>|<span data-ttu-id="894db-284">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="894db-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="894db-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="894db-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="894db-286">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="894db-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-287">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-287">Requirements</span></span>

|<span data-ttu-id="894db-288">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-288">Requirement</span></span>| <span data-ttu-id="894db-289">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-290">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-291">1.3</span><span class="sxs-lookup"><span data-stu-id="894db-291">1.3</span></span>|
|[<span data-ttu-id="894db-292">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-293">Restreinte</span><span class="sxs-lookup"><span data-stu-id="894db-293">Restricted</span></span>|
|[<span data-ttu-id="894db-294">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-295">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="894db-296">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="894db-296">Returns:</span></span>

<span data-ttu-id="894db-297">Type : String</span><span class="sxs-lookup"><span data-stu-id="894db-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="894db-298">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="894db-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="894db-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="894db-300">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="894db-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="894db-301">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="894db-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-302">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-302">Parameters</span></span>

|<span data-ttu-id="894db-303">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-303">Name</span></span>| <span data-ttu-id="894db-304">Type</span><span class="sxs-lookup"><span data-stu-id="894db-304">Type</span></span>| <span data-ttu-id="894db-305">Description</span><span class="sxs-lookup"><span data-stu-id="894db-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="894db-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="894db-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="894db-307">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="894db-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-308">Requirements</span></span>

|<span data-ttu-id="894db-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-309">Requirement</span></span>| <span data-ttu-id="894db-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-312">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-312">1.0</span></span>|
|[<span data-ttu-id="894db-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-314">ReadItem</span></span>|
|[<span data-ttu-id="894db-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-316">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="894db-317">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="894db-317">Returns:</span></span>

<span data-ttu-id="894db-318">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="894db-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="894db-319">Type : Date</span><span class="sxs-lookup"><span data-stu-id="894db-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="894db-320">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="894db-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="894db-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="894db-322">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="894db-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-323">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="894db-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="894db-324">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="894db-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="894db-p111">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="894db-p111">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="894db-327">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="894db-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="894db-328">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="894db-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-329">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-329">Parameters</span></span>

|<span data-ttu-id="894db-330">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-330">Name</span></span>| <span data-ttu-id="894db-331">Type</span><span class="sxs-lookup"><span data-stu-id="894db-331">Type</span></span>| <span data-ttu-id="894db-332">Description</span><span class="sxs-lookup"><span data-stu-id="894db-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="894db-333">String</span><span class="sxs-lookup"><span data-stu-id="894db-333">String</span></span>|<span data-ttu-id="894db-334">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="894db-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-335">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-335">Requirements</span></span>

|<span data-ttu-id="894db-336">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-336">Requirement</span></span>| <span data-ttu-id="894db-337">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-338">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-339">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-339">1.0</span></span>|
|[<span data-ttu-id="894db-340">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-341">ReadItem</span></span>|
|[<span data-ttu-id="894db-342">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-343">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="894db-344">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="894db-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="894db-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="894db-346">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="894db-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-347">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="894db-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="894db-348">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="894db-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="894db-349">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="894db-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="894db-350">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="894db-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="894db-p112">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="894db-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-353">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-353">Parameters</span></span>

|<span data-ttu-id="894db-354">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-354">Name</span></span>| <span data-ttu-id="894db-355">Type</span><span class="sxs-lookup"><span data-stu-id="894db-355">Type</span></span>| <span data-ttu-id="894db-356">Description</span><span class="sxs-lookup"><span data-stu-id="894db-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="894db-357">String</span><span class="sxs-lookup"><span data-stu-id="894db-357">String</span></span>|<span data-ttu-id="894db-358">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="894db-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-359">Requirements</span></span>

|<span data-ttu-id="894db-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-360">Requirement</span></span>| <span data-ttu-id="894db-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-363">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-363">1.0</span></span>|
|[<span data-ttu-id="894db-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-365">ReadItem</span></span>|
|[<span data-ttu-id="894db-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-367">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="894db-368">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="894db-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="894db-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="894db-370">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="894db-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-371">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="894db-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="894db-p113">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="894db-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="894db-p114">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="894db-p114">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="894db-p115">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="894db-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="894db-379">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="894db-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-380">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-380">Parameters</span></span>

|<span data-ttu-id="894db-381">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-381">Name</span></span>| <span data-ttu-id="894db-382">Type</span><span class="sxs-lookup"><span data-stu-id="894db-382">Type</span></span>| <span data-ttu-id="894db-383">Description</span><span class="sxs-lookup"><span data-stu-id="894db-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="894db-384">Object</span><span class="sxs-lookup"><span data-stu-id="894db-384">Object</span></span> | <span data-ttu-id="894db-385">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="894db-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="894db-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="894db-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="894db-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="894db-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="894db-p117">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="894db-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="894db-392">Date</span><span class="sxs-lookup"><span data-stu-id="894db-392">Date</span></span> | <span data-ttu-id="894db-393">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="894db-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="894db-394">Date</span><span class="sxs-lookup"><span data-stu-id="894db-394">Date</span></span> | <span data-ttu-id="894db-395">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="894db-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="894db-396">String</span><span class="sxs-lookup"><span data-stu-id="894db-396">String</span></span> | <span data-ttu-id="894db-p118">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="894db-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="894db-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="894db-p119">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="894db-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="894db-402">String</span><span class="sxs-lookup"><span data-stu-id="894db-402">String</span></span> | <span data-ttu-id="894db-p120">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="894db-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="894db-405">String</span><span class="sxs-lookup"><span data-stu-id="894db-405">String</span></span> | <span data-ttu-id="894db-p121">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="894db-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="894db-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-408">Requirements</span></span>

|<span data-ttu-id="894db-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-409">Requirement</span></span>| <span data-ttu-id="894db-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-412">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-412">1.0</span></span>|
|[<span data-ttu-id="894db-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-414">ReadItem</span></span>|
|[<span data-ttu-id="894db-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="894db-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="894db-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="894db-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="894db-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="894db-419">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="894db-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="894db-p122">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="894db-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-422">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="894db-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="894db-423">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="894db-423">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="894db-424">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="894db-424">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="894db-425">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="894db-425">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="894db-426">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="894db-426">**REST Tokens**</span></span>

<span data-ttu-id="894db-p124">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="894db-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="894db-430">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="894db-430">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="894db-431">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="894db-431">**EWS Tokens**</span></span>

<span data-ttu-id="894db-p125">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="894db-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="894db-434">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="894db-434">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="894db-435">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="894db-435">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="894db-436">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="894db-436">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="894db-437">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="894db-437">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-438">Parameters</span><span class="sxs-lookup"><span data-stu-id="894db-438">Parameters</span></span>

|<span data-ttu-id="894db-439">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-439">Name</span></span>| <span data-ttu-id="894db-440">Type</span><span class="sxs-lookup"><span data-stu-id="894db-440">Type</span></span>| <span data-ttu-id="894db-441">Attributs</span><span class="sxs-lookup"><span data-stu-id="894db-441">Attributes</span></span>| <span data-ttu-id="894db-442">Description</span><span class="sxs-lookup"><span data-stu-id="894db-442">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="894db-443">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-443">Object</span></span> | <span data-ttu-id="894db-444">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-444">&lt;optional&gt;</span></span> | <span data-ttu-id="894db-445">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="894db-445">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="894db-446">Boolean</span><span class="sxs-lookup"><span data-stu-id="894db-446">Boolean</span></span> |  <span data-ttu-id="894db-447">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-447">&lt;optional&gt;</span></span> | <span data-ttu-id="894db-p127">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="894db-p127">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="894db-450">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-450">Object</span></span> |  <span data-ttu-id="894db-451">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-451">&lt;optional&gt;</span></span> | <span data-ttu-id="894db-452">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="894db-452">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="894db-453">fonction</span><span class="sxs-lookup"><span data-stu-id="894db-453">function</span></span>||<span data-ttu-id="894db-454">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="894db-454">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="894db-455">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="894db-455">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="894db-456">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="894db-456">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="894db-457">Erreurs</span><span class="sxs-lookup"><span data-stu-id="894db-457">Errors</span></span>

|<span data-ttu-id="894db-458">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="894db-458">Error code</span></span>|<span data-ttu-id="894db-459">Description</span><span class="sxs-lookup"><span data-stu-id="894db-459">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="894db-460">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="894db-460">The request has failed.</span></span> <span data-ttu-id="894db-461">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="894db-461">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="894db-462">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="894db-462">The Exchange server returned an error.</span></span> <span data-ttu-id="894db-463">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="894db-463">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="894db-464">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="894db-464">The user is no longer connected to the network.</span></span> <span data-ttu-id="894db-465">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="894db-465">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-466">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-466">Requirements</span></span>

|<span data-ttu-id="894db-467">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-467">Requirement</span></span>| <span data-ttu-id="894db-468">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-469">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-470">1,5</span><span class="sxs-lookup"><span data-stu-id="894db-470">1.5</span></span> |
|[<span data-ttu-id="894db-471">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-472">ReadItem</span></span>|
|[<span data-ttu-id="894db-473">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-474">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="894db-474">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="894db-475">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-475">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="894db-476">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="894db-476">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="894db-477">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="894db-477">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="894db-p131">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="894db-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="894db-480">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="894db-480">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="894db-481">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="894db-481">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="894db-482">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="894db-482">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="894db-483">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="894db-483">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="894db-484">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="894db-484">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="894db-485">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="894db-485">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-486">Parameters</span><span class="sxs-lookup"><span data-stu-id="894db-486">Parameters</span></span>

|<span data-ttu-id="894db-487">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-487">Name</span></span>| <span data-ttu-id="894db-488">Type</span><span class="sxs-lookup"><span data-stu-id="894db-488">Type</span></span>| <span data-ttu-id="894db-489">Attributs</span><span class="sxs-lookup"><span data-stu-id="894db-489">Attributes</span></span>| <span data-ttu-id="894db-490">Description</span><span class="sxs-lookup"><span data-stu-id="894db-490">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="894db-491">function</span><span class="sxs-lookup"><span data-stu-id="894db-491">function</span></span>||<span data-ttu-id="894db-492">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="894db-492">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="894db-493">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="894db-493">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="894db-494">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="894db-494">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="894db-495">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-495">Object</span></span>| <span data-ttu-id="894db-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-496">&lt;optional&gt;</span></span>|<span data-ttu-id="894db-497">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="894db-497">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="894db-498">Erreurs</span><span class="sxs-lookup"><span data-stu-id="894db-498">Errors</span></span>

|<span data-ttu-id="894db-499">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="894db-499">Error code</span></span>|<span data-ttu-id="894db-500">Description</span><span class="sxs-lookup"><span data-stu-id="894db-500">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="894db-501">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="894db-501">The request has failed.</span></span> <span data-ttu-id="894db-502">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="894db-502">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="894db-503">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="894db-503">The Exchange server returned an error.</span></span> <span data-ttu-id="894db-504">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="894db-504">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="894db-505">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="894db-505">The user is no longer connected to the network.</span></span> <span data-ttu-id="894db-506">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="894db-506">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-507">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-507">Requirements</span></span>

|<span data-ttu-id="894db-508">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-508">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="894db-509">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-510">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-510">1.0</span></span> | <span data-ttu-id="894db-511">1.3</span><span class="sxs-lookup"><span data-stu-id="894db-511">1.3</span></span> |
|[<span data-ttu-id="894db-512">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-513">ReadItem</span></span> | <span data-ttu-id="894db-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-514">ReadItem</span></span> |
|[<span data-ttu-id="894db-515">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-515">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-516">Lecture</span><span class="sxs-lookup"><span data-stu-id="894db-516">Read</span></span> | <span data-ttu-id="894db-517">Composition</span><span class="sxs-lookup"><span data-stu-id="894db-517">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="894db-518">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-518">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="894db-519">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="894db-519">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="894db-520">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="894db-520">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="894db-521">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="894db-521">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-522">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-522">Parameters</span></span>

|<span data-ttu-id="894db-523">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-523">Name</span></span>| <span data-ttu-id="894db-524">Type</span><span class="sxs-lookup"><span data-stu-id="894db-524">Type</span></span>| <span data-ttu-id="894db-525">Attributs</span><span class="sxs-lookup"><span data-stu-id="894db-525">Attributes</span></span>| <span data-ttu-id="894db-526">Description</span><span class="sxs-lookup"><span data-stu-id="894db-526">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="894db-527">function</span><span class="sxs-lookup"><span data-stu-id="894db-527">function</span></span>||<span data-ttu-id="894db-528">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="894db-528">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="894db-529">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="894db-529">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="894db-530">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="894db-530">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="894db-531">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-531">Object</span></span>| <span data-ttu-id="894db-532">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-532">&lt;optional&gt;</span></span>|<span data-ttu-id="894db-533">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="894db-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="894db-534">Erreurs</span><span class="sxs-lookup"><span data-stu-id="894db-534">Errors</span></span>

|<span data-ttu-id="894db-535">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="894db-535">Error code</span></span>|<span data-ttu-id="894db-536">Description</span><span class="sxs-lookup"><span data-stu-id="894db-536">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="894db-537">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="894db-537">The request has failed.</span></span> <span data-ttu-id="894db-538">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="894db-538">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="894db-539">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="894db-539">The Exchange server returned an error.</span></span> <span data-ttu-id="894db-540">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="894db-540">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="894db-541">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="894db-541">The user is no longer connected to the network.</span></span> <span data-ttu-id="894db-542">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="894db-542">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-543">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-543">Requirements</span></span>

|<span data-ttu-id="894db-544">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-544">Requirement</span></span>| <span data-ttu-id="894db-545">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-546">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-547">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-547">1.0</span></span>|
|[<span data-ttu-id="894db-548">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-549">ReadItem</span></span>|
|[<span data-ttu-id="894db-550">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-551">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="894db-552">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-552">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="894db-553">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="894db-553">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="894db-554">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="894db-554">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="894db-555">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="894db-555">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="894db-556">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="894db-556">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="894db-557">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="894db-557">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="894db-558">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="894db-558">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="894db-559">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="894db-559">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="894db-560">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="894db-560">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="894db-561">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="894db-561">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="894db-562">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="894db-562">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="894db-p141">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="894db-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="894db-565">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="894db-565">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="894db-566">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="894db-566">Version differences</span></span>

<span data-ttu-id="894db-567">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="894db-567">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="894db-p142">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="894db-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-571">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-571">Parameters</span></span>

|<span data-ttu-id="894db-572">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-572">Name</span></span>| <span data-ttu-id="894db-573">Type</span><span class="sxs-lookup"><span data-stu-id="894db-573">Type</span></span>| <span data-ttu-id="894db-574">Attributs</span><span class="sxs-lookup"><span data-stu-id="894db-574">Attributes</span></span>| <span data-ttu-id="894db-575">Description</span><span class="sxs-lookup"><span data-stu-id="894db-575">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="894db-576">String</span><span class="sxs-lookup"><span data-stu-id="894db-576">String</span></span>||<span data-ttu-id="894db-577">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="894db-577">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="894db-578">function</span><span class="sxs-lookup"><span data-stu-id="894db-578">function</span></span>||<span data-ttu-id="894db-579">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="894db-579">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="894db-580">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="894db-580">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="894db-581">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="894db-581">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="894db-582">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-582">Object</span></span>| <span data-ttu-id="894db-583">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-583">&lt;optional&gt;</span></span>|<span data-ttu-id="894db-584">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="894db-584">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-585">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-585">Requirements</span></span>

|<span data-ttu-id="894db-586">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-586">Requirement</span></span>| <span data-ttu-id="894db-587">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-588">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-589">1.0</span><span class="sxs-lookup"><span data-stu-id="894db-589">1.0</span></span>|
|[<span data-ttu-id="894db-590">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-591">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="894db-591">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="894db-592">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-593">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-593">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="894db-594">Exemple</span><span class="sxs-lookup"><span data-stu-id="894db-594">Example</span></span>

<span data-ttu-id="894db-595">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="894db-595">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="894db-596">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="894db-596">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="894db-597">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="894db-597">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="894db-598">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="894db-598">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="894db-599">Paramètres</span><span class="sxs-lookup"><span data-stu-id="894db-599">Parameters</span></span>

| <span data-ttu-id="894db-600">Nom</span><span class="sxs-lookup"><span data-stu-id="894db-600">Name</span></span> | <span data-ttu-id="894db-601">Type</span><span class="sxs-lookup"><span data-stu-id="894db-601">Type</span></span> | <span data-ttu-id="894db-602">Attributs</span><span class="sxs-lookup"><span data-stu-id="894db-602">Attributes</span></span> | <span data-ttu-id="894db-603">Description</span><span class="sxs-lookup"><span data-stu-id="894db-603">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="894db-604">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="894db-604">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="894db-605">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="894db-605">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="894db-606">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-606">Object</span></span> | <span data-ttu-id="894db-607">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-607">&lt;optional&gt;</span></span> | <span data-ttu-id="894db-608">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="894db-608">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="894db-609">Objet</span><span class="sxs-lookup"><span data-stu-id="894db-609">Object</span></span> | <span data-ttu-id="894db-610">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-610">&lt;optional&gt;</span></span> | <span data-ttu-id="894db-611">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="894db-611">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="894db-612">fonction</span><span class="sxs-lookup"><span data-stu-id="894db-612">function</span></span>| <span data-ttu-id="894db-613">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="894db-613">&lt;optional&gt;</span></span>|<span data-ttu-id="894db-614">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="894db-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="894db-615">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="894db-615">Requirements</span></span>

|<span data-ttu-id="894db-616">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="894db-616">Requirement</span></span>| <span data-ttu-id="894db-617">Valeur</span><span class="sxs-lookup"><span data-stu-id="894db-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="894db-618">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="894db-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="894db-619">1,5</span><span class="sxs-lookup"><span data-stu-id="894db-619">1.5</span></span> |
|[<span data-ttu-id="894db-620">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="894db-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="894db-621">ReadItem</span><span class="sxs-lookup"><span data-stu-id="894db-621">ReadItem</span></span> |
|[<span data-ttu-id="894db-622">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="894db-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="894db-623">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="894db-623">Compose or Read</span></span>|
