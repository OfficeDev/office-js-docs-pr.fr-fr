---
title: Office.context-ensemble de conditions requises 1.6
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: 336357d5915a6b061e69ef488eb31a11077722b1
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/13/2019
ms.locfileid: "29389563"
---
# <a name="mailbox"></a><span data-ttu-id="48ade-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="48ade-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="48ade-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="48ade-104">Permet d’accéder au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="48ade-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="48ade-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-105">Requirements</span></span>

|<span data-ttu-id="48ade-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-106">Requirement</span></span>| <span data-ttu-id="48ade-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-109">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-109">1.0</span></span>|
|[<span data-ttu-id="48ade-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="48ade-111">Restricted</span></span>|
|[<span data-ttu-id="48ade-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="48ade-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="48ade-114">Members and methods</span></span>

| <span data-ttu-id="48ade-115">Membre</span><span class="sxs-lookup"><span data-stu-id="48ade-115">Member</span></span> | <span data-ttu-id="48ade-116">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="48ade-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="48ade-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="48ade-118">Membre</span><span class="sxs-lookup"><span data-stu-id="48ade-118">Member</span></span> |
| [<span data-ttu-id="48ade-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="48ade-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="48ade-120">Membre</span><span class="sxs-lookup"><span data-stu-id="48ade-120">Member</span></span> |
| [<span data-ttu-id="48ade-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="48ade-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="48ade-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-122">Method</span></span> |
| [<span data-ttu-id="48ade-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="48ade-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="48ade-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-124">Method</span></span> |
| [<span data-ttu-id="48ade-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="48ade-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="48ade-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-126">Method</span></span> |
| [<span data-ttu-id="48ade-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="48ade-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="48ade-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-128">Method</span></span> |
| [<span data-ttu-id="48ade-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="48ade-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="48ade-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-130">Method</span></span> |
| [<span data-ttu-id="48ade-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="48ade-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="48ade-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-132">Method</span></span> |
| [<span data-ttu-id="48ade-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="48ade-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="48ade-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-134">Method</span></span> |
| [<span data-ttu-id="48ade-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="48ade-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="48ade-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-136">Method</span></span> |
| [<span data-ttu-id="48ade-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="48ade-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="48ade-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-138">Method</span></span> |
| [<span data-ttu-id="48ade-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="48ade-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="48ade-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-140">Method</span></span> |
| [<span data-ttu-id="48ade-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="48ade-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="48ade-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-142">Method</span></span> |
| [<span data-ttu-id="48ade-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="48ade-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="48ade-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-144">Method</span></span> |
| [<span data-ttu-id="48ade-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="48ade-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="48ade-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-146">Method</span></span> |
| [<span data-ttu-id="48ade-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="48ade-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="48ade-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="48ade-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="48ade-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="48ade-149">Namespaces</span></span>

<span data-ttu-id="48ade-150">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="48ade-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="48ade-151">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="48ade-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="48ade-152">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="48ade-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="48ade-153">Membres</span><span class="sxs-lookup"><span data-stu-id="48ade-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="48ade-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="48ade-154">ewsUrl :String</span></span>

<span data-ttu-id="48ade-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="48ade-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-157">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="48ade-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="48ade-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="48ade-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="48ade-160">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="48ade-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="48ade-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="48ade-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="48ade-163">Type :</span><span class="sxs-lookup"><span data-stu-id="48ade-163">Type:</span></span>

*   <span data-ttu-id="48ade-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48ade-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-165">Requirements</span></span>

|<span data-ttu-id="48ade-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-166">Requirement</span></span>| <span data-ttu-id="48ade-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-169">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-169">1.0</span></span>|
|[<span data-ttu-id="48ade-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-171">ReadItem</span></span>|
|[<span data-ttu-id="48ade-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-173">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-173">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="48ade-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="48ade-174">restUrl :String</span></span>

<span data-ttu-id="48ade-175">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="48ade-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="48ade-176">La valeur `restUrl` peut être utilisée pour que l’[API REST](https://docs.microsoft.com/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="48ade-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="48ade-177">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="48ade-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="48ade-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="48ade-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="48ade-180">Type :</span><span class="sxs-lookup"><span data-stu-id="48ade-180">Type:</span></span>

*   <span data-ttu-id="48ade-181">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48ade-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-182">Requirements</span></span>

|<span data-ttu-id="48ade-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-183">Requirement</span></span>| <span data-ttu-id="48ade-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-186">1,5</span><span class="sxs-lookup"><span data-stu-id="48ade-186">1.5</span></span> |
|[<span data-ttu-id="48ade-187">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-188">ReadItem</span></span>|
|[<span data-ttu-id="48ade-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-190">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-190">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="48ade-191">Méthodes</span><span class="sxs-lookup"><span data-stu-id="48ade-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="48ade-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48ade-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="48ade-193">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="48ade-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="48ade-194">Actuellement, le seul type d’événement pris en charge est `Office.EventType.ItemChanged`, qui est appelé quand l’utilisateur sélectionne un nouvel élément.</span><span class="sxs-lookup"><span data-stu-id="48ade-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="48ade-195">Cet événement est utilisé par les compléments qui implémentent un volet Office épinglable. Il les autorise à actualiser l’IU du volet Office à partir de l’élément sélectionné.</span><span class="sxs-lookup"><span data-stu-id="48ade-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-196">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-196">Parameters:</span></span>

| <span data-ttu-id="48ade-197">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-197">Name</span></span> | <span data-ttu-id="48ade-198">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-198">Type</span></span> | <span data-ttu-id="48ade-199">Attributs</span><span class="sxs-lookup"><span data-stu-id="48ade-199">Attributes</span></span> | <span data-ttu-id="48ade-200">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="48ade-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="48ade-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="48ade-202">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="48ade-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="48ade-203">Fonction</span><span class="sxs-lookup"><span data-stu-id="48ade-203">Function</span></span> || <span data-ttu-id="48ade-p106">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="48ade-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="48ade-207">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-207">Object</span></span> | <span data-ttu-id="48ade-208">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-208">&lt;optional&gt;</span></span> | <span data-ttu-id="48ade-209">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="48ade-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="48ade-210">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-210">Object</span></span> | <span data-ttu-id="48ade-211">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-211">&lt;optional&gt;</span></span> | <span data-ttu-id="48ade-212">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="48ade-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="48ade-213">fonction</span><span class="sxs-lookup"><span data-stu-id="48ade-213">function</span></span>| <span data-ttu-id="48ade-214">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-214">&lt;optional&gt;</span></span>|<span data-ttu-id="48ade-215">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48ade-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-216">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-216">Requirements</span></span>

|<span data-ttu-id="48ade-217">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-217">Requirement</span></span>| <span data-ttu-id="48ade-218">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-219">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-220">1,5</span><span class="sxs-lookup"><span data-stu-id="48ade-220">1.5</span></span> |
|[<span data-ttu-id="48ade-221">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-222">ReadItem</span></span> |
|[<span data-ttu-id="48ade-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-224">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-224">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-225">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="48ade-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="48ade-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="48ade-227">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="48ade-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-228">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="48ade-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="48ade-p107">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="48ade-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-231">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-231">Parameters:</span></span>

|<span data-ttu-id="48ade-232">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-232">Name</span></span>| <span data-ttu-id="48ade-233">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-233">Type</span></span>| <span data-ttu-id="48ade-234">object</span><span class="sxs-lookup"><span data-stu-id="48ade-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="48ade-235">String</span><span class="sxs-lookup"><span data-stu-id="48ade-235">String</span></span>|<span data-ttu-id="48ade-236">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="48ade-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="48ade-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="48ade-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="48ade-238">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="48ade-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-239">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-239">Requirements</span></span>

|<span data-ttu-id="48ade-240">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-240">Requirement</span></span>| <span data-ttu-id="48ade-241">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-242">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-243">1.3</span><span class="sxs-lookup"><span data-stu-id="48ade-243">1.3</span></span>|
|[<span data-ttu-id="48ade-244">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-244">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-245">Restreinte</span><span class="sxs-lookup"><span data-stu-id="48ade-245">Restricted</span></span>|
|[<span data-ttu-id="48ade-246">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-246">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-247">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-247">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48ade-248">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="48ade-248">Returns:</span></span>

<span data-ttu-id="48ade-249">Type : String</span><span class="sxs-lookup"><span data-stu-id="48ade-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="48ade-250">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-250">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="48ade-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="48ade-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="48ade-252">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="48ade-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="48ade-p108">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="48ade-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="48ade-p109">Si l’application de messagerie est en cours d’exécution dans Outlook, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook Web App, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="48ade-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-258">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-258">Parameters:</span></span>

|<span data-ttu-id="48ade-259">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-259">Name</span></span>| <span data-ttu-id="48ade-260">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-260">Type</span></span>| <span data-ttu-id="48ade-261">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="48ade-262">Date</span><span class="sxs-lookup"><span data-stu-id="48ade-262">Date</span></span>|<span data-ttu-id="48ade-263">Objet Date</span><span class="sxs-lookup"><span data-stu-id="48ade-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-264">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-264">Requirements</span></span>

|<span data-ttu-id="48ade-265">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-265">Requirement</span></span>| <span data-ttu-id="48ade-266">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-267">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-268">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-268">1.0</span></span>|
|[<span data-ttu-id="48ade-269">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-270">ReadItem</span></span>|
|[<span data-ttu-id="48ade-271">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-272">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-272">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48ade-273">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="48ade-273">Returns:</span></span>

<span data-ttu-id="48ade-274">Type : [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="48ade-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="48ade-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="48ade-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="48ade-276">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="48ade-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-277">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="48ade-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="48ade-p110">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="48ade-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-280">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-280">Parameters:</span></span>

|<span data-ttu-id="48ade-281">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-281">Name</span></span>| <span data-ttu-id="48ade-282">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-282">Type</span></span>| <span data-ttu-id="48ade-283">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="48ade-284">String</span><span class="sxs-lookup"><span data-stu-id="48ade-284">String</span></span>|<span data-ttu-id="48ade-285">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="48ade-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="48ade-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="48ade-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="48ade-287">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="48ade-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-288">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-288">Requirements</span></span>

|<span data-ttu-id="48ade-289">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-289">Requirement</span></span>| <span data-ttu-id="48ade-290">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-291">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-292">1.3</span><span class="sxs-lookup"><span data-stu-id="48ade-292">1.3</span></span>|
|[<span data-ttu-id="48ade-293">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-294">Restreinte</span><span class="sxs-lookup"><span data-stu-id="48ade-294">Restricted</span></span>|
|[<span data-ttu-id="48ade-295">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-296">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-296">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48ade-297">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="48ade-297">Returns:</span></span>

<span data-ttu-id="48ade-298">Type : String</span><span class="sxs-lookup"><span data-stu-id="48ade-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="48ade-299">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-299">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="48ade-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="48ade-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="48ade-301">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="48ade-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="48ade-302">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="48ade-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-303">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-303">Parameters:</span></span>

|<span data-ttu-id="48ade-304">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-304">Name</span></span>| <span data-ttu-id="48ade-305">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-305">Type</span></span>| <span data-ttu-id="48ade-306">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="48ade-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="48ade-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="48ade-308">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="48ade-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-309">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-309">Requirements</span></span>

|<span data-ttu-id="48ade-310">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-310">Requirement</span></span>| <span data-ttu-id="48ade-311">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-312">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-313">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-313">1.0</span></span>|
|[<span data-ttu-id="48ade-314">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-315">ReadItem</span></span>|
|[<span data-ttu-id="48ade-316">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-317">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-317">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="48ade-318">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="48ade-318">Returns:</span></span>

<span data-ttu-id="48ade-319">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="48ade-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="48ade-320">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="48ade-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="48ade-321">Date</span><span class="sxs-lookup"><span data-stu-id="48ade-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="48ade-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="48ade-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="48ade-323">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="48ade-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-324">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="48ade-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="48ade-325">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="48ade-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="48ade-p111">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="48ade-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="48ade-328">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="48ade-329">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="48ade-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-330">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-330">Parameters:</span></span>

|<span data-ttu-id="48ade-331">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-331">Name</span></span>| <span data-ttu-id="48ade-332">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-332">Type</span></span>| <span data-ttu-id="48ade-333">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="48ade-334">String</span><span class="sxs-lookup"><span data-stu-id="48ade-334">String</span></span>|<span data-ttu-id="48ade-335">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="48ade-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-336">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-336">Requirements</span></span>

|<span data-ttu-id="48ade-337">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-337">Requirement</span></span>| <span data-ttu-id="48ade-338">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-339">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-340">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-340">1.0</span></span>|
|[<span data-ttu-id="48ade-341">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-341">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-342">ReadItem</span></span>|
|[<span data-ttu-id="48ade-343">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-343">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-344">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-344">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-345">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-345">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="48ade-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="48ade-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="48ade-347">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="48ade-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-348">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="48ade-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="48ade-349">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="48ade-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="48ade-350">Dans Outlook Web App, cette méthode ouvre le formulaire indiqué uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="48ade-351">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="48ade-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="48ade-p112">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="48ade-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-354">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-354">Parameters:</span></span>

|<span data-ttu-id="48ade-355">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-355">Name</span></span>| <span data-ttu-id="48ade-356">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-356">Type</span></span>| <span data-ttu-id="48ade-357">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="48ade-358">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-358">String</span></span>|<span data-ttu-id="48ade-359">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="48ade-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-360">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-360">Requirements</span></span>

|<span data-ttu-id="48ade-361">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-361">Requirement</span></span>| <span data-ttu-id="48ade-362">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-363">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-364">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-364">1.0</span></span>|
|[<span data-ttu-id="48ade-365">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-365">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-366">ReadItem</span></span>|
|[<span data-ttu-id="48ade-367">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-368">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-368">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-369">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-369">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="48ade-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="48ade-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="48ade-371">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="48ade-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-372">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="48ade-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="48ade-p113">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="48ade-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="48ade-p114">Dans Outlook Web App et OWA pour les périphériques, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="48ade-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="48ade-p115">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="48ade-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="48ade-380">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="48ade-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-381">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-381">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-382">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="48ade-382">All parameters are optional.</span></span>

|<span data-ttu-id="48ade-383">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-383">Name</span></span>| <span data-ttu-id="48ade-384">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-384">Type</span></span>| <span data-ttu-id="48ade-385">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="48ade-386">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-386">Object</span></span> | <span data-ttu-id="48ade-387">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="48ade-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="48ade-388">Tableau.&lt;Chaîne&gt; &#124; Tableau.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="48ade-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="48ade-391">Tableau.&lt;Chaîne&gt; &#124; Tableau.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="48ade-p117">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="48ade-394">Date</span><span class="sxs-lookup"><span data-stu-id="48ade-394">Date</span></span> | <span data-ttu-id="48ade-395">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="48ade-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="48ade-396">Date</span><span class="sxs-lookup"><span data-stu-id="48ade-396">Date</span></span> | <span data-ttu-id="48ade-397">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="48ade-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="48ade-398">String</span><span class="sxs-lookup"><span data-stu-id="48ade-398">String</span></span> | <span data-ttu-id="48ade-p118">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="48ade-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="48ade-p119">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="48ade-404">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-404">String</span></span> | <span data-ttu-id="48ade-p120">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="48ade-407">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-407">String</span></span> | <span data-ttu-id="48ade-p121">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="48ade-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="48ade-410">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-410">Requirements</span></span>

|<span data-ttu-id="48ade-411">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-411">Requirement</span></span>| <span data-ttu-id="48ade-412">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-413">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-414">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-414">1.0</span></span>|
|[<span data-ttu-id="48ade-415">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-415">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-416">ReadItem</span></span>|
|[<span data-ttu-id="48ade-417">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-417">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-418">Lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-419">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="48ade-420">displayNewMessageForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="48ade-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="48ade-421">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="48ade-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="48ade-422">La méthode `displayNewMessageForm` ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="48ade-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="48ade-423">Si des paramètres sont spécifiés, les champs du formulaire de message sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="48ade-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="48ade-424">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="48ade-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-425">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-425">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-426">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="48ade-426">All parameters are optional.</span></span>

|<span data-ttu-id="48ade-427">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-427">Name</span></span>| <span data-ttu-id="48ade-428">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-428">Type</span></span>| <span data-ttu-id="48ade-429">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="48ade-430">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-430">Object</span></span> | <span data-ttu-id="48ade-431">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="48ade-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="48ade-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="48ade-433">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne À.</span><span class="sxs-lookup"><span data-stu-id="48ade-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="48ade-434">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="48ade-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="48ade-436">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cc.</span><span class="sxs-lookup"><span data-stu-id="48ade-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="48ade-437">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="48ade-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="48ade-439">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cci.</span><span class="sxs-lookup"><span data-stu-id="48ade-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="48ade-440">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="48ade-441">String</span><span class="sxs-lookup"><span data-stu-id="48ade-441">String</span></span> | <span data-ttu-id="48ade-442">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="48ade-442">A string containing the subject of the message.</span></span> <span data-ttu-id="48ade-443">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="48ade-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="48ade-444">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-444">String</span></span> | <span data-ttu-id="48ade-445">Corps du message HTML.</span><span class="sxs-lookup"><span data-stu-id="48ade-445">The HTML body of the message.</span></span> <span data-ttu-id="48ade-446">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="48ade-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="48ade-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="48ade-448">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="48ade-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="48ade-449">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-449">String</span></span> | <span data-ttu-id="48ade-p128">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="48ade-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="48ade-452">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-452">String</span></span> | <span data-ttu-id="48ade-453">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="48ade-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="48ade-454">Chaîne</span><span class="sxs-lookup"><span data-stu-id="48ade-454">String</span></span> | <span data-ttu-id="48ade-p129">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="48ade-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="48ade-457">Boolean</span><span class="sxs-lookup"><span data-stu-id="48ade-457">Boolean</span></span> | <span data-ttu-id="48ade-p130">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="48ade-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="48ade-460">String</span><span class="sxs-lookup"><span data-stu-id="48ade-460">String</span></span> | <span data-ttu-id="48ade-461">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="48ade-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="48ade-462">ID d’élément EWS du courrier électronique existant à joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="48ade-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="48ade-463">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="48ade-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="48ade-464">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-464">Requirements</span></span>

|<span data-ttu-id="48ade-465">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-465">Requirement</span></span>| <span data-ttu-id="48ade-466">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-467">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-468">1.6</span><span class="sxs-lookup"><span data-stu-id="48ade-468">1.6</span></span> |
|[<span data-ttu-id="48ade-469">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-470">ReadItem</span></span>|
|[<span data-ttu-id="48ade-471">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-472">Lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-473">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-473">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="48ade-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="48ade-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="48ade-475">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="48ade-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="48ade-p132">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="48ade-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-478">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="48ade-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="48ade-479">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="48ade-479">**REST Tokens**</span></span>

<span data-ttu-id="48ade-p133">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="48ade-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="48ade-483">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="48ade-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="48ade-484">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="48ade-484">**EWS Tokens**</span></span>

<span data-ttu-id="48ade-p134">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="48ade-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="48ade-487">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="48ade-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-488">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-488">Parameters:</span></span>

|<span data-ttu-id="48ade-489">Name</span><span class="sxs-lookup"><span data-stu-id="48ade-489">Name</span></span>| <span data-ttu-id="48ade-490">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-490">Type</span></span>| <span data-ttu-id="48ade-491">Attributs</span><span class="sxs-lookup"><span data-stu-id="48ade-491">Attributes</span></span>| <span data-ttu-id="48ade-492">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="48ade-493">Object</span><span class="sxs-lookup"><span data-stu-id="48ade-493">Object</span></span> | <span data-ttu-id="48ade-494">&lt;facultatif&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-494">&lt;optional&gt;</span></span> | <span data-ttu-id="48ade-495">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="48ade-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="48ade-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="48ade-496">Boolean</span></span> |  <span data-ttu-id="48ade-497">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-497">&lt;optional&gt;</span></span> | <span data-ttu-id="48ade-p135">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="48ade-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="48ade-500">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-500">Object</span></span> |  <span data-ttu-id="48ade-501">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-501">&lt;optional&gt;</span></span> | <span data-ttu-id="48ade-502">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="48ade-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="48ade-503">fonction</span><span class="sxs-lookup"><span data-stu-id="48ade-503">function</span></span>||<span data-ttu-id="48ade-p136">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48ade-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-506">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-506">Requirements</span></span>

|<span data-ttu-id="48ade-507">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-507">Requirement</span></span>| <span data-ttu-id="48ade-508">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-509">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-510">1,5</span><span class="sxs-lookup"><span data-stu-id="48ade-510">1.5</span></span> |
|[<span data-ttu-id="48ade-511">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-511">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-512">ReadItem</span></span>|
|[<span data-ttu-id="48ade-513">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-513">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-514">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-515">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="48ade-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="48ade-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="48ade-517">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="48ade-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="48ade-p137">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="48ade-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="48ade-p138">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="48ade-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="48ade-523">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="48ade-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="48ade-p139">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="48ade-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-526">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-526">Parameters:</span></span>

|<span data-ttu-id="48ade-527">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-527">Name</span></span>| <span data-ttu-id="48ade-528">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-528">Type</span></span>| <span data-ttu-id="48ade-529">Attributs</span><span class="sxs-lookup"><span data-stu-id="48ade-529">Attributes</span></span>| <span data-ttu-id="48ade-530">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="48ade-531">fonction</span><span class="sxs-lookup"><span data-stu-id="48ade-531">function</span></span>||<span data-ttu-id="48ade-p140">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48ade-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="48ade-534">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-534">Object</span></span>| <span data-ttu-id="48ade-535">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-535">&lt;optional&gt;</span></span>|<span data-ttu-id="48ade-536">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="48ade-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-537">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-537">Requirements</span></span>

|<span data-ttu-id="48ade-538">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-538">Requirement</span></span>| <span data-ttu-id="48ade-539">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-540">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-541">1.3</span><span class="sxs-lookup"><span data-stu-id="48ade-541">1.3</span></span>|
|[<span data-ttu-id="48ade-542">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-543">ReadItem</span></span>|
|[<span data-ttu-id="48ade-544">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-545">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-546">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-546">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="48ade-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="48ade-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="48ade-548">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="48ade-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="48ade-549">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="48ade-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-550">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-550">Parameters:</span></span>

|<span data-ttu-id="48ade-551">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-551">Name</span></span>| <span data-ttu-id="48ade-552">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-552">Type</span></span>| <span data-ttu-id="48ade-553">Attributs</span><span class="sxs-lookup"><span data-stu-id="48ade-553">Attributes</span></span>| <span data-ttu-id="48ade-554">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="48ade-555">function</span><span class="sxs-lookup"><span data-stu-id="48ade-555">function</span></span>||<span data-ttu-id="48ade-556">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48ade-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48ade-557">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48ade-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="48ade-558">Object</span><span class="sxs-lookup"><span data-stu-id="48ade-558">Object</span></span>| <span data-ttu-id="48ade-559">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-559">&lt;optional&gt;</span></span>|<span data-ttu-id="48ade-560">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="48ade-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-561">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-561">Requirements</span></span>

|<span data-ttu-id="48ade-562">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-562">Requirement</span></span>| <span data-ttu-id="48ade-563">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-564">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-565">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-565">1.0</span></span>|
|[<span data-ttu-id="48ade-566">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-567">ReadItem</span></span>|
|[<span data-ttu-id="48ade-568">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-569">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-569">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-570">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="48ade-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="48ade-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="48ade-572">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="48ade-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-573">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="48ade-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="48ade-574">dans Outlook pour iOS ou Outlook pour Android ;</span><span class="sxs-lookup"><span data-stu-id="48ade-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="48ade-575">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="48ade-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="48ade-576">Dans ces cas de figure, les compléments doivent [utiliser les API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="48ade-576">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="48ade-577">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="48ade-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="48ade-578">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="48ade-578">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="48ade-579">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="48ade-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="48ade-580">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="48ade-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="48ade-p142">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="48ade-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="48ade-583">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="48ade-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="48ade-584">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="48ade-584">Version differences</span></span>

<span data-ttu-id="48ade-585">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="48ade-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="48ade-p143">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="48ade-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-589">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-589">Parameters:</span></span>

|<span data-ttu-id="48ade-590">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-590">Name</span></span>| <span data-ttu-id="48ade-591">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-591">Type</span></span>| <span data-ttu-id="48ade-592">Attributs</span><span class="sxs-lookup"><span data-stu-id="48ade-592">Attributes</span></span>| <span data-ttu-id="48ade-593">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="48ade-594">String</span><span class="sxs-lookup"><span data-stu-id="48ade-594">String</span></span>||<span data-ttu-id="48ade-595">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="48ade-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="48ade-596">function</span><span class="sxs-lookup"><span data-stu-id="48ade-596">function</span></span>||<span data-ttu-id="48ade-597">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48ade-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="48ade-598">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="48ade-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="48ade-599">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="48ade-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="48ade-600">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-600">Object</span></span>| <span data-ttu-id="48ade-601">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-601">&lt;optional&gt;</span></span>|<span data-ttu-id="48ade-602">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="48ade-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-603">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-603">Requirements</span></span>

|<span data-ttu-id="48ade-604">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-604">Requirement</span></span>| <span data-ttu-id="48ade-605">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-606">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-607">1.0</span><span class="sxs-lookup"><span data-stu-id="48ade-607">1.0</span></span>|
|[<span data-ttu-id="48ade-608">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="48ade-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="48ade-610">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-611">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-611">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="48ade-612">Exemple</span><span class="sxs-lookup"><span data-stu-id="48ade-612">Example</span></span>

<span data-ttu-id="48ade-613">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="48ade-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="48ade-614">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="48ade-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="48ade-615">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="48ade-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="48ade-616">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="48ade-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="48ade-617">Paramètres :</span><span class="sxs-lookup"><span data-stu-id="48ade-617">Parameters:</span></span>

| <span data-ttu-id="48ade-618">Nom</span><span class="sxs-lookup"><span data-stu-id="48ade-618">Name</span></span> | <span data-ttu-id="48ade-619">Type</span><span class="sxs-lookup"><span data-stu-id="48ade-619">Type</span></span> | <span data-ttu-id="48ade-620">Attributs</span><span class="sxs-lookup"><span data-stu-id="48ade-620">Attributes</span></span> | <span data-ttu-id="48ade-621">Description</span><span class="sxs-lookup"><span data-stu-id="48ade-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="48ade-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="48ade-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="48ade-623">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="48ade-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="48ade-624">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-624">Object</span></span> | <span data-ttu-id="48ade-625">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-625">&lt;optional&gt;</span></span> | <span data-ttu-id="48ade-626">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="48ade-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="48ade-627">Objet</span><span class="sxs-lookup"><span data-stu-id="48ade-627">Object</span></span> | <span data-ttu-id="48ade-628">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-628">&lt;optional&gt;</span></span> | <span data-ttu-id="48ade-629">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="48ade-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="48ade-630">fonction</span><span class="sxs-lookup"><span data-stu-id="48ade-630">function</span></span>| <span data-ttu-id="48ade-631">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="48ade-631">&lt;optional&gt;</span></span>|<span data-ttu-id="48ade-632">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="48ade-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="48ade-633">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48ade-633">Requirements</span></span>

|<span data-ttu-id="48ade-634">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48ade-634">Requirement</span></span>| <span data-ttu-id="48ade-635">Valeur</span><span class="sxs-lookup"><span data-stu-id="48ade-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="48ade-636">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48ade-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48ade-637">1,5</span><span class="sxs-lookup"><span data-stu-id="48ade-637">1.5</span></span> |
|[<span data-ttu-id="48ade-638">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48ade-638">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48ade-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="48ade-639">ReadItem</span></span> |
|[<span data-ttu-id="48ade-640">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48ade-640">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="48ade-641">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="48ade-641">Compose or read</span></span>|
