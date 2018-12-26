---
title: Office.context – ensemble de conditions requises 1.5
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: c80ed3837315f3bf51da302d91f08e2114af3b2f
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433956"
---
# <a name="mailbox"></a><span data-ttu-id="45e83-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="45e83-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="45e83-103">Office.context.mailbox</span></span>

<span data-ttu-id="45e83-104">Permet d’accéder au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="45e83-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="45e83-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-105">Requirements</span></span>

|<span data-ttu-id="45e83-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-106">Requirement</span></span>| <span data-ttu-id="45e83-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-109">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-109">1.0</span></span>|
|[<span data-ttu-id="45e83-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="45e83-111">Restricted</span></span>|
|[<span data-ttu-id="45e83-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="45e83-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="45e83-114">Members and methods</span></span>

| <span data-ttu-id="45e83-115">Membre</span><span class="sxs-lookup"><span data-stu-id="45e83-115">Member</span></span> | <span data-ttu-id="45e83-116">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="45e83-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="45e83-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="45e83-118">Membre</span><span class="sxs-lookup"><span data-stu-id="45e83-118">Member</span></span> |
| [<span data-ttu-id="45e83-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="45e83-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="45e83-120">Membre</span><span class="sxs-lookup"><span data-stu-id="45e83-120">Member</span></span> |
| [<span data-ttu-id="45e83-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="45e83-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="45e83-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-122">Method</span></span> |
| [<span data-ttu-id="45e83-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="45e83-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="45e83-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-124">Method</span></span> |
| [<span data-ttu-id="45e83-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="45e83-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime) | <span data-ttu-id="45e83-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-126">Method</span></span> |
| [<span data-ttu-id="45e83-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="45e83-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="45e83-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-128">Method</span></span> |
| [<span data-ttu-id="45e83-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="45e83-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="45e83-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-130">Method</span></span> |
| [<span data-ttu-id="45e83-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="45e83-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="45e83-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-132">Method</span></span> |
| [<span data-ttu-id="45e83-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="45e83-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="45e83-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-134">Method</span></span> |
| [<span data-ttu-id="45e83-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="45e83-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="45e83-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-136">Method</span></span> |
| [<span data-ttu-id="45e83-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="45e83-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="45e83-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-138">Method</span></span> |
| [<span data-ttu-id="45e83-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="45e83-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="45e83-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-140">Method</span></span> |
| [<span data-ttu-id="45e83-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="45e83-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="45e83-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-142">Method</span></span> |
| [<span data-ttu-id="45e83-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="45e83-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="45e83-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-144">Method</span></span> |
| [<span data-ttu-id="45e83-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="45e83-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="45e83-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="45e83-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="45e83-147">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="45e83-147">Namespaces</span></span>

<span data-ttu-id="45e83-148">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="45e83-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="45e83-149">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="45e83-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="45e83-150">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="45e83-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="45e83-151">Membres</span><span class="sxs-lookup"><span data-stu-id="45e83-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="45e83-152">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="45e83-152">ewsUrl :String</span></span>

<span data-ttu-id="45e83-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="45e83-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-155">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="45e83-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45e83-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="45e83-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="45e83-158">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="45e83-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="45e83-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="45e83-161">Type :</span><span class="sxs-lookup"><span data-stu-id="45e83-161">Type:</span></span>

*   <span data-ttu-id="45e83-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="45e83-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45e83-163">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-163">Requirements</span></span>

|<span data-ttu-id="45e83-164">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-164">Requirement</span></span>| <span data-ttu-id="45e83-165">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-166">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-167">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-167">1.0</span></span>|
|[<span data-ttu-id="45e83-168">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-168">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-169">ReadItem</span></span>|
|[<span data-ttu-id="45e83-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-171">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-171">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="45e83-172">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="45e83-172">restUrl :String</span></span>

<span data-ttu-id="45e83-173">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="45e83-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="45e83-174">La valeur `restUrl` peut être utilisée pour que l’[API REST](https://docs.microsoft.com/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="45e83-174">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="45e83-175">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="45e83-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="45e83-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-178">Les clients Outlook connectés aux installations locales d’Exchange 2016 ou version ultérieure avec une URL REST personnalisée configurée renvoient une valeur non valide pour `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="45e83-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="45e83-179">Type :</span><span class="sxs-lookup"><span data-stu-id="45e83-179">Type:</span></span>

*   <span data-ttu-id="45e83-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="45e83-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="45e83-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-181">Requirements</span></span>

|<span data-ttu-id="45e83-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-182">Requirement</span></span>| <span data-ttu-id="45e83-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-185">1,5</span><span class="sxs-lookup"><span data-stu-id="45e83-185">1.5</span></span> |
|[<span data-ttu-id="45e83-186">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-186">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-187">ReadItem</span></span>|
|[<span data-ttu-id="45e83-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-188">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-189">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-189">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="45e83-190">Méthodes</span><span class="sxs-lookup"><span data-stu-id="45e83-190">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="45e83-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45e83-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="45e83-192">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="45e83-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="45e83-193">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé quand l’utilisateur sélectionne un nouvel élément.</span><span class="sxs-lookup"><span data-stu-id="45e83-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="45e83-194">Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’IU du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="45e83-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-195">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-195">Parameters:</span></span>

| <span data-ttu-id="45e83-196">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-196">Name</span></span> | <span data-ttu-id="45e83-197">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-197">Type</span></span> | <span data-ttu-id="45e83-198">Attributs</span><span class="sxs-lookup"><span data-stu-id="45e83-198">Attributes</span></span> | <span data-ttu-id="45e83-199">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="45e83-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="45e83-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="45e83-201">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="45e83-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="45e83-202">Fonction</span><span class="sxs-lookup"><span data-stu-id="45e83-202">Function</span></span> || <span data-ttu-id="45e83-p106">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="45e83-206">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-206">Object</span></span> | <span data-ttu-id="45e83-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-207">&lt;optional&gt;</span></span> | <span data-ttu-id="45e83-208">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="45e83-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="45e83-209">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-209">Object</span></span> | <span data-ttu-id="45e83-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-210">&lt;optional&gt;</span></span> | <span data-ttu-id="45e83-211">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="45e83-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="45e83-212">fonction</span><span class="sxs-lookup"><span data-stu-id="45e83-212">function</span></span>| <span data-ttu-id="45e83-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-213">&lt;optional&gt;</span></span>|<span data-ttu-id="45e83-214">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45e83-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-215">Requirements</span></span>

|<span data-ttu-id="45e83-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-216">Requirement</span></span>| <span data-ttu-id="45e83-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-219">1,5</span><span class="sxs-lookup"><span data-stu-id="45e83-219">1.5</span></span> |
|[<span data-ttu-id="45e83-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-221">ReadItem</span></span> |
|[<span data-ttu-id="45e83-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-223">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-223">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-224">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="45e83-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="45e83-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="45e83-226">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="45e83-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-227">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="45e83-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45e83-p107">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="45e83-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-230">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-230">Parameters:</span></span>

|<span data-ttu-id="45e83-231">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-231">Name</span></span>| <span data-ttu-id="45e83-232">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-232">Type</span></span>| <span data-ttu-id="45e83-233">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="45e83-234">String</span><span class="sxs-lookup"><span data-stu-id="45e83-234">String</span></span>|<span data-ttu-id="45e83-235">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="45e83-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="45e83-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="45e83-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="45e83-237">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="45e83-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-238">Requirements</span></span>

|<span data-ttu-id="45e83-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-239">Requirement</span></span>| <span data-ttu-id="45e83-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-242">1.3</span><span class="sxs-lookup"><span data-stu-id="45e83-242">1.3</span></span>|
|[<span data-ttu-id="45e83-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-244">Restreinte</span><span class="sxs-lookup"><span data-stu-id="45e83-244">Restricted</span></span>|
|[<span data-ttu-id="45e83-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-246">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-246">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45e83-247">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="45e83-247">Returns:</span></span>

<span data-ttu-id="45e83-248">Type : String</span><span class="sxs-lookup"><span data-stu-id="45e83-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="45e83-249">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-249">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="45e83-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="45e83-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="45e83-251">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="45e83-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="45e83-p108">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="45e83-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="45e83-p109">Si l’application de messagerie est en cours d’exécution dans Outlook, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook Web App, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="45e83-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-257">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-257">Parameters:</span></span>

|<span data-ttu-id="45e83-258">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-258">Name</span></span>| <span data-ttu-id="45e83-259">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-259">Type</span></span>| <span data-ttu-id="45e83-260">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="45e83-261">Date</span><span class="sxs-lookup"><span data-stu-id="45e83-261">Date</span></span>|<span data-ttu-id="45e83-262">Objet Date</span><span class="sxs-lookup"><span data-stu-id="45e83-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-263">Requirements</span></span>

|<span data-ttu-id="45e83-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-264">Requirement</span></span>| <span data-ttu-id="45e83-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-267">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-267">1.0</span></span>|
|[<span data-ttu-id="45e83-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-269">ReadItem</span></span>|
|[<span data-ttu-id="45e83-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-271">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-271">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45e83-272">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="45e83-272">Returns:</span></span>

<span data-ttu-id="45e83-273">Type : [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="45e83-273">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="45e83-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="45e83-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="45e83-275">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="45e83-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-276">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="45e83-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45e83-p110">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="45e83-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-279">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-279">Parameters:</span></span>

|<span data-ttu-id="45e83-280">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-280">Name</span></span>| <span data-ttu-id="45e83-281">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-281">Type</span></span>| <span data-ttu-id="45e83-282">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="45e83-283">String</span><span class="sxs-lookup"><span data-stu-id="45e83-283">String</span></span>|<span data-ttu-id="45e83-284">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="45e83-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="45e83-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="45e83-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="45e83-286">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="45e83-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-287">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-287">Requirements</span></span>

|<span data-ttu-id="45e83-288">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-288">Requirement</span></span>| <span data-ttu-id="45e83-289">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-290">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-291">1.3</span><span class="sxs-lookup"><span data-stu-id="45e83-291">1.3</span></span>|
|[<span data-ttu-id="45e83-292">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-292">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-293">Restreinte</span><span class="sxs-lookup"><span data-stu-id="45e83-293">Restricted</span></span>|
|[<span data-ttu-id="45e83-294">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-294">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-295">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-295">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45e83-296">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="45e83-296">Returns:</span></span>

<span data-ttu-id="45e83-297">Type : String</span><span class="sxs-lookup"><span data-stu-id="45e83-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="45e83-298">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-298">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="45e83-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="45e83-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="45e83-300">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="45e83-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="45e83-301">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="45e83-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-302">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-302">Parameters:</span></span>

|<span data-ttu-id="45e83-303">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-303">Name</span></span>| <span data-ttu-id="45e83-304">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-304">Type</span></span>| <span data-ttu-id="45e83-305">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="45e83-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="45e83-306">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="45e83-307">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="45e83-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-308">Requirements</span></span>

|<span data-ttu-id="45e83-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-309">Requirement</span></span>| <span data-ttu-id="45e83-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-312">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-312">1.0</span></span>|
|[<span data-ttu-id="45e83-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-314">ReadItem</span></span>|
|[<span data-ttu-id="45e83-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-316">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-316">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="45e83-317">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="45e83-317">Returns:</span></span>

<span data-ttu-id="45e83-318">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="45e83-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="45e83-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="45e83-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="45e83-320">Date</span><span class="sxs-lookup"><span data-stu-id="45e83-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="45e83-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="45e83-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="45e83-322">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="45e83-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-323">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="45e83-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45e83-324">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="45e83-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="45e83-p111">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="45e83-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="45e83-327">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="45e83-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="45e83-328">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="45e83-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-329">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-329">Parameters:</span></span>

|<span data-ttu-id="45e83-330">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-330">Name</span></span>| <span data-ttu-id="45e83-331">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-331">Type</span></span>| <span data-ttu-id="45e83-332">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="45e83-333">String</span><span class="sxs-lookup"><span data-stu-id="45e83-333">String</span></span>|<span data-ttu-id="45e83-334">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="45e83-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-335">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-335">Requirements</span></span>

|<span data-ttu-id="45e83-336">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-336">Requirement</span></span>| <span data-ttu-id="45e83-337">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-338">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-339">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-339">1.0</span></span>|
|[<span data-ttu-id="45e83-340">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-340">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-341">ReadItem</span></span>|
|[<span data-ttu-id="45e83-342">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-342">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-343">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-343">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-344">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="45e83-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="45e83-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="45e83-346">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="45e83-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-347">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="45e83-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45e83-348">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="45e83-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="45e83-349">Dans Outlook Web App, cette méthode ouvre le formulaire indiqué uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="45e83-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="45e83-350">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="45e83-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="45e83-p112">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="45e83-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-353">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-353">Parameters:</span></span>

|<span data-ttu-id="45e83-354">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-354">Name</span></span>| <span data-ttu-id="45e83-355">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-355">Type</span></span>| <span data-ttu-id="45e83-356">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="45e83-357">String</span><span class="sxs-lookup"><span data-stu-id="45e83-357">String</span></span>|<span data-ttu-id="45e83-358">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="45e83-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-359">Requirements</span></span>

|<span data-ttu-id="45e83-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-360">Requirement</span></span>| <span data-ttu-id="45e83-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-363">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-363">1.0</span></span>|
|[<span data-ttu-id="45e83-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-365">ReadItem</span></span>|
|[<span data-ttu-id="45e83-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-367">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-367">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-368">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="45e83-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="45e83-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="45e83-370">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="45e83-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-371">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="45e83-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="45e83-p113">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="45e83-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="45e83-p114">Dans Outlook Web App et OWA pour les périphériques, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="45e83-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="45e83-p115">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="45e83-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="45e83-379">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="45e83-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-380">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-380">Parameters:</span></span>

|<span data-ttu-id="45e83-381">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-381">Name</span></span>| <span data-ttu-id="45e83-382">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-382">Type</span></span>| <span data-ttu-id="45e83-383">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="45e83-384">Object</span><span class="sxs-lookup"><span data-stu-id="45e83-384">Object</span></span> | <span data-ttu-id="45e83-385">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="45e83-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="45e83-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="45e83-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="45e83-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="45e83-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="45e83-p117">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="45e83-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="45e83-392">Date</span><span class="sxs-lookup"><span data-stu-id="45e83-392">Date</span></span> | <span data-ttu-id="45e83-393">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="45e83-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="45e83-394">Date</span><span class="sxs-lookup"><span data-stu-id="45e83-394">Date</span></span> | <span data-ttu-id="45e83-395">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="45e83-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="45e83-396">String</span><span class="sxs-lookup"><span data-stu-id="45e83-396">String</span></span> | <span data-ttu-id="45e83-p118">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="45e83-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="45e83-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="45e83-p119">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="45e83-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="45e83-402">String</span><span class="sxs-lookup"><span data-stu-id="45e83-402">String</span></span> | <span data-ttu-id="45e83-p120">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="45e83-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="45e83-405">String</span><span class="sxs-lookup"><span data-stu-id="45e83-405">String</span></span> | <span data-ttu-id="45e83-p121">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="45e83-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="45e83-408">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-408">Requirements</span></span>

|<span data-ttu-id="45e83-409">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-409">Requirement</span></span>| <span data-ttu-id="45e83-410">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-411">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-412">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-412">1.0</span></span>|
|[<span data-ttu-id="45e83-413">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-413">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-414">ReadItem</span></span>|
|[<span data-ttu-id="45e83-415">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-415">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-416">Lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-417">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="45e83-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="45e83-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="45e83-419">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="45e83-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="45e83-p122">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="45e83-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-422">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="45e83-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="45e83-423">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="45e83-423">**REST Tokens**</span></span>

<span data-ttu-id="45e83-p123">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="45e83-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="45e83-427">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="45e83-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="45e83-428">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="45e83-428">**EWS Tokens**</span></span>

<span data-ttu-id="45e83-p124">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="45e83-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="45e83-431">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="45e83-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-432">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-432">Parameters:</span></span>

|<span data-ttu-id="45e83-433">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-433">Name</span></span>| <span data-ttu-id="45e83-434">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-434">Type</span></span>| <span data-ttu-id="45e83-435">Attributs</span><span class="sxs-lookup"><span data-stu-id="45e83-435">Attributes</span></span>| <span data-ttu-id="45e83-436">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="45e83-437">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-437">Object</span></span> | <span data-ttu-id="45e83-438">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-438">&lt;optional&gt;</span></span> | <span data-ttu-id="45e83-439">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="45e83-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="45e83-440">Boolean</span><span class="sxs-lookup"><span data-stu-id="45e83-440">Boolean</span></span> |  <span data-ttu-id="45e83-441">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-441">&lt;optional&gt;</span></span> | <span data-ttu-id="45e83-p125">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="45e83-444">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-444">Object</span></span> |  <span data-ttu-id="45e83-445">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-445">&lt;optional&gt;</span></span> | <span data-ttu-id="45e83-446">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="45e83-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="45e83-447">fonction</span><span class="sxs-lookup"><span data-stu-id="45e83-447">function</span></span>||<span data-ttu-id="45e83-p126">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p126">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-450">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-450">Requirements</span></span>

|<span data-ttu-id="45e83-451">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-451">Requirement</span></span>| <span data-ttu-id="45e83-452">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-453">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-454">1,5</span><span class="sxs-lookup"><span data-stu-id="45e83-454">1.5</span></span> |
|[<span data-ttu-id="45e83-455">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-455">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-456">ReadItem</span></span>|
|[<span data-ttu-id="45e83-457">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-457">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-458">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-458">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-459">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-459">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="45e83-460">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="45e83-460">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="45e83-461">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="45e83-461">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="45e83-p127">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="45e83-p127">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="45e83-p128">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="45e83-p128">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="45e83-467">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="45e83-467">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="45e83-p129">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p129">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-470">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-470">Parameters:</span></span>

|<span data-ttu-id="45e83-471">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-471">Name</span></span>| <span data-ttu-id="45e83-472">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-472">Type</span></span>| <span data-ttu-id="45e83-473">Attributs</span><span class="sxs-lookup"><span data-stu-id="45e83-473">Attributes</span></span>| <span data-ttu-id="45e83-474">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-474">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="45e83-475">function</span><span class="sxs-lookup"><span data-stu-id="45e83-475">function</span></span>||<span data-ttu-id="45e83-p130">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p130">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="45e83-478">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-478">Object</span></span>| <span data-ttu-id="45e83-479">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-479">&lt;optional&gt;</span></span>|<span data-ttu-id="45e83-480">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="45e83-480">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-481">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-481">Requirements</span></span>

|<span data-ttu-id="45e83-482">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-482">Requirement</span></span>| <span data-ttu-id="45e83-483">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-484">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-485">1.3</span><span class="sxs-lookup"><span data-stu-id="45e83-485">1.3</span></span>|
|[<span data-ttu-id="45e83-486">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-486">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-487">ReadItem</span></span>|
|[<span data-ttu-id="45e83-488">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-488">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-489">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-489">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-490">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-490">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="45e83-491">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="45e83-491">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="45e83-492">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="45e83-492">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="45e83-493">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="45e83-493">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-494">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-494">Parameters:</span></span>

|<span data-ttu-id="45e83-495">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-495">Name</span></span>| <span data-ttu-id="45e83-496">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-496">Type</span></span>| <span data-ttu-id="45e83-497">Attributs</span><span class="sxs-lookup"><span data-stu-id="45e83-497">Attributes</span></span>| <span data-ttu-id="45e83-498">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-498">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="45e83-499">function</span><span class="sxs-lookup"><span data-stu-id="45e83-499">function</span></span>||<span data-ttu-id="45e83-500">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45e83-500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45e83-501">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45e83-501">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="45e83-502">Object</span><span class="sxs-lookup"><span data-stu-id="45e83-502">Object</span></span>| <span data-ttu-id="45e83-503">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-503">&lt;optional&gt;</span></span>|<span data-ttu-id="45e83-504">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="45e83-504">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-505">Requirements</span></span>

|<span data-ttu-id="45e83-506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-506">Requirement</span></span>| <span data-ttu-id="45e83-507">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-509">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-509">1.0</span></span>|
|[<span data-ttu-id="45e83-510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-511">ReadItem</span></span>|
|[<span data-ttu-id="45e83-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-513">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-513">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-514">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-514">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="45e83-515">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="45e83-515">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="45e83-516">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="45e83-516">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-517">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="45e83-517">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="45e83-518">dans Outlook pour iOS ou Outlook pour Android ;</span><span class="sxs-lookup"><span data-stu-id="45e83-518">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="45e83-519">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="45e83-519">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="45e83-520">Dans ces cas de figure, les compléments doivent [utiliser les API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="45e83-520">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="45e83-521">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="45e83-521">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="45e83-522">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="45e83-522">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="45e83-523">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="45e83-523">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="45e83-524">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="45e83-524">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="45e83-p132">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="45e83-p132">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="45e83-527">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="45e83-527">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="45e83-528">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="45e83-528">Version differences</span></span>

<span data-ttu-id="45e83-529">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="45e83-529">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="45e83-p133">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="45e83-p133">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-533">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-533">Parameters:</span></span>

|<span data-ttu-id="45e83-534">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-534">Name</span></span>| <span data-ttu-id="45e83-535">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-535">Type</span></span>| <span data-ttu-id="45e83-536">Attributs</span><span class="sxs-lookup"><span data-stu-id="45e83-536">Attributes</span></span>| <span data-ttu-id="45e83-537">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-537">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="45e83-538">String</span><span class="sxs-lookup"><span data-stu-id="45e83-538">String</span></span>||<span data-ttu-id="45e83-539">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="45e83-539">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="45e83-540">function</span><span class="sxs-lookup"><span data-stu-id="45e83-540">function</span></span>||<span data-ttu-id="45e83-541">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45e83-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="45e83-542">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="45e83-542">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="45e83-543">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="45e83-543">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="45e83-544">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-544">Object</span></span>| <span data-ttu-id="45e83-545">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-545">&lt;optional&gt;</span></span>|<span data-ttu-id="45e83-546">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="45e83-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-547">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-547">Requirements</span></span>

|<span data-ttu-id="45e83-548">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-548">Requirement</span></span>| <span data-ttu-id="45e83-549">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-550">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-551">1.0</span><span class="sxs-lookup"><span data-stu-id="45e83-551">1.0</span></span>|
|[<span data-ttu-id="45e83-552">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-552">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-553">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="45e83-553">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="45e83-554">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-554">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-555">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-555">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="45e83-556">Exemple</span><span class="sxs-lookup"><span data-stu-id="45e83-556">Example</span></span>

<span data-ttu-id="45e83-557">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="45e83-557">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="45e83-558">removeHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="45e83-558">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="45e83-559">Retire un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="45e83-559">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="45e83-560">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="45e83-560">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="45e83-561">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="45e83-561">Parameters:</span></span>

| <span data-ttu-id="45e83-562">Nom</span><span class="sxs-lookup"><span data-stu-id="45e83-562">Name</span></span> | <span data-ttu-id="45e83-563">Type</span><span class="sxs-lookup"><span data-stu-id="45e83-563">Type</span></span> | <span data-ttu-id="45e83-564">Attributs</span><span class="sxs-lookup"><span data-stu-id="45e83-564">Attributes</span></span> | <span data-ttu-id="45e83-565">Description</span><span class="sxs-lookup"><span data-stu-id="45e83-565">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="45e83-566">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="45e83-566">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="45e83-567">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="45e83-567">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="45e83-568">Fonction</span><span class="sxs-lookup"><span data-stu-id="45e83-568">Function</span></span> || <span data-ttu-id="45e83-p135">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="45e83-p135">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="45e83-572">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-572">Object</span></span> | <span data-ttu-id="45e83-573">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-573">&lt;optional&gt;</span></span> | <span data-ttu-id="45e83-574">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="45e83-574">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="45e83-575">Objet</span><span class="sxs-lookup"><span data-stu-id="45e83-575">Object</span></span> | <span data-ttu-id="45e83-576">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-576">&lt;optional&gt;</span></span> | <span data-ttu-id="45e83-577">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="45e83-577">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="45e83-578">fonction</span><span class="sxs-lookup"><span data-stu-id="45e83-578">function</span></span>| <span data-ttu-id="45e83-579">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="45e83-579">&lt;optional&gt;</span></span>|<span data-ttu-id="45e83-580">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="45e83-580">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="45e83-581">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="45e83-581">Requirements</span></span>

|<span data-ttu-id="45e83-582">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="45e83-582">Requirement</span></span>| <span data-ttu-id="45e83-583">Valeur</span><span class="sxs-lookup"><span data-stu-id="45e83-583">Value</span></span>|
|---|---|
|[<span data-ttu-id="45e83-584">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="45e83-584">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="45e83-585">1,5</span><span class="sxs-lookup"><span data-stu-id="45e83-585">1.5</span></span> |
|[<span data-ttu-id="45e83-586">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="45e83-586">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="45e83-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="45e83-587">ReadItem</span></span> |
|[<span data-ttu-id="45e83-588">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="45e83-588">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="45e83-589">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="45e83-589">Compose or read</span></span>|