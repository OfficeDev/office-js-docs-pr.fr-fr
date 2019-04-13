---
title: Office. Context. Mailbox-ensemble de conditions requises 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 0f84e657644b198fbca722a0628a5bafcce84377
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838549"
---
# <a name="mailbox"></a><span data-ttu-id="725b8-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="725b8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="725b8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="725b8-104">Permet d’accéder au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="725b8-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="725b8-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-105">Requirements</span></span>

|<span data-ttu-id="725b8-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-106">Requirement</span></span>| <span data-ttu-id="725b8-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-109">1.0</span></span>|
|[<span data-ttu-id="725b8-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="725b8-111">Restricted</span></span>|
|[<span data-ttu-id="725b8-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="725b8-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="725b8-114">Members and methods</span></span>

| <span data-ttu-id="725b8-115">Membre</span><span class="sxs-lookup"><span data-stu-id="725b8-115">Member</span></span> | <span data-ttu-id="725b8-116">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="725b8-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="725b8-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="725b8-118">Membre</span><span class="sxs-lookup"><span data-stu-id="725b8-118">Member</span></span> |
| [<span data-ttu-id="725b8-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="725b8-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="725b8-120">Membre</span><span class="sxs-lookup"><span data-stu-id="725b8-120">Member</span></span> |
| [<span data-ttu-id="725b8-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="725b8-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="725b8-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-122">Method</span></span> |
| [<span data-ttu-id="725b8-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="725b8-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="725b8-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-124">Method</span></span> |
| [<span data-ttu-id="725b8-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="725b8-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="725b8-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-126">Method</span></span> |
| [<span data-ttu-id="725b8-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="725b8-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="725b8-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-128">Method</span></span> |
| [<span data-ttu-id="725b8-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="725b8-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="725b8-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-130">Method</span></span> |
| [<span data-ttu-id="725b8-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="725b8-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="725b8-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-132">Method</span></span> |
| [<span data-ttu-id="725b8-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="725b8-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="725b8-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-134">Method</span></span> |
| [<span data-ttu-id="725b8-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="725b8-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="725b8-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-136">Method</span></span> |
| [<span data-ttu-id="725b8-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="725b8-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="725b8-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-138">Method</span></span> |
| [<span data-ttu-id="725b8-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="725b8-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="725b8-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-140">Method</span></span> |
| [<span data-ttu-id="725b8-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="725b8-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="725b8-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-142">Method</span></span> |
| [<span data-ttu-id="725b8-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="725b8-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="725b8-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-144">Method</span></span> |
| [<span data-ttu-id="725b8-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="725b8-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="725b8-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-146">Method</span></span> |
| [<span data-ttu-id="725b8-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="725b8-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="725b8-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="725b8-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="725b8-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="725b8-149">Namespaces</span></span>

<span data-ttu-id="725b8-150">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="725b8-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="725b8-151">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="725b8-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="725b8-152">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="725b8-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="725b8-153">Membres</span><span class="sxs-lookup"><span data-stu-id="725b8-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="725b8-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="725b8-154">ewsUrl :String</span></span>

<span data-ttu-id="725b8-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="725b8-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-157">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="725b8-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="725b8-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="725b8-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="725b8-160">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="725b8-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="725b8-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="725b8-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="725b8-163">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-163">Type</span></span>

*   <span data-ttu-id="725b8-164">String</span><span class="sxs-lookup"><span data-stu-id="725b8-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="725b8-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-165">Requirements</span></span>

|<span data-ttu-id="725b8-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-166">Requirement</span></span>| <span data-ttu-id="725b8-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-169">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-169">1.0</span></span>|
|[<span data-ttu-id="725b8-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-171">ReadItem</span></span>|
|[<span data-ttu-id="725b8-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-173">Compose or Read</span></span>|

---
---

#### <a name="resturl-string"></a><span data-ttu-id="725b8-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="725b8-174">restUrl :String</span></span>

<span data-ttu-id="725b8-175">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="725b8-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="725b8-176">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="725b8-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="725b8-177">L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.</span><span class="sxs-lookup"><span data-stu-id="725b8-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="725b8-p104">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="725b8-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="725b8-180">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-180">Type</span></span>

*   <span data-ttu-id="725b8-181">String</span><span class="sxs-lookup"><span data-stu-id="725b8-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="725b8-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-182">Requirements</span></span>

|<span data-ttu-id="725b8-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-183">Requirement</span></span>| <span data-ttu-id="725b8-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-186">1,5</span><span class="sxs-lookup"><span data-stu-id="725b8-186">1.5</span></span> |
|[<span data-ttu-id="725b8-187">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-187">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-188">ReadItem</span></span>|
|[<span data-ttu-id="725b8-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-189">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="725b8-191">Méthodes</span><span class="sxs-lookup"><span data-stu-id="725b8-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="725b8-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="725b8-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="725b8-193">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="725b8-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="725b8-194">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="725b8-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-195">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-195">Parameters</span></span>

| <span data-ttu-id="725b8-196">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-196">Name</span></span> | <span data-ttu-id="725b8-197">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-197">Type</span></span> | <span data-ttu-id="725b8-198">Attributs</span><span class="sxs-lookup"><span data-stu-id="725b8-198">Attributes</span></span> | <span data-ttu-id="725b8-199">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="725b8-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="725b8-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="725b8-201">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="725b8-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="725b8-202">Fonction</span><span class="sxs-lookup"><span data-stu-id="725b8-202">Function</span></span> || <span data-ttu-id="725b8-p105">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="725b8-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="725b8-206">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-206">Object</span></span> | <span data-ttu-id="725b8-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-207">&lt;optional&gt;</span></span> | <span data-ttu-id="725b8-208">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="725b8-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="725b8-209">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-209">Object</span></span> | <span data-ttu-id="725b8-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-210">&lt;optional&gt;</span></span> | <span data-ttu-id="725b8-211">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="725b8-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="725b8-212">fonction</span><span class="sxs-lookup"><span data-stu-id="725b8-212">function</span></span>| <span data-ttu-id="725b8-213">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-213">&lt;optional&gt;</span></span>|<span data-ttu-id="725b8-214">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="725b8-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-215">Requirements</span></span>

|<span data-ttu-id="725b8-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-216">Requirement</span></span>| <span data-ttu-id="725b8-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-219">1,5</span><span class="sxs-lookup"><span data-stu-id="725b8-219">1.5</span></span> |
|[<span data-ttu-id="725b8-220">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-221">ReadItem</span></span> |
|[<span data-ttu-id="725b8-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-223">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-224">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-224">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="725b8-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="725b8-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="725b8-226">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="725b8-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-227">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="725b8-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="725b8-p106">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="725b8-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-230">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-230">Parameters</span></span>

|<span data-ttu-id="725b8-231">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-231">Name</span></span>| <span data-ttu-id="725b8-232">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-232">Type</span></span>| <span data-ttu-id="725b8-233">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="725b8-234">String</span><span class="sxs-lookup"><span data-stu-id="725b8-234">String</span></span>|<span data-ttu-id="725b8-235">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="725b8-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="725b8-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="725b8-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="725b8-237">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="725b8-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-238">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-238">Requirements</span></span>

|<span data-ttu-id="725b8-239">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-239">Requirement</span></span>| <span data-ttu-id="725b8-240">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-241">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-242">1.3</span><span class="sxs-lookup"><span data-stu-id="725b8-242">1.3</span></span>|
|[<span data-ttu-id="725b8-243">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-244">Restreinte</span><span class="sxs-lookup"><span data-stu-id="725b8-244">Restricted</span></span>|
|[<span data-ttu-id="725b8-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="725b8-247">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="725b8-247">Returns:</span></span>

<span data-ttu-id="725b8-248">Type : String</span><span class="sxs-lookup"><span data-stu-id="725b8-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="725b8-249">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="725b8-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="725b8-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="725b8-251">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="725b8-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="725b8-p107">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="725b8-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="725b8-p108">Si l’application de messagerie est en cours d’exécution dans Outlook, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook Web App, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="725b8-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-257">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-257">Parameters</span></span>

|<span data-ttu-id="725b8-258">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-258">Name</span></span>| <span data-ttu-id="725b8-259">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-259">Type</span></span>| <span data-ttu-id="725b8-260">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="725b8-261">Date</span><span class="sxs-lookup"><span data-stu-id="725b8-261">Date</span></span>|<span data-ttu-id="725b8-262">Objet Date</span><span class="sxs-lookup"><span data-stu-id="725b8-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-263">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-263">Requirements</span></span>

|<span data-ttu-id="725b8-264">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-264">Requirement</span></span>| <span data-ttu-id="725b8-265">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-266">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-267">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-267">1.0</span></span>|
|[<span data-ttu-id="725b8-268">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-269">ReadItem</span></span>|
|[<span data-ttu-id="725b8-270">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-271">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="725b8-272">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="725b8-272">Returns:</span></span>

<span data-ttu-id="725b8-273">Type : [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="725b8-273">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

---
---

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="725b8-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="725b8-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="725b8-275">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="725b8-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-276">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="725b8-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="725b8-p109">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="725b8-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-279">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-279">Parameters</span></span>

|<span data-ttu-id="725b8-280">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-280">Name</span></span>| <span data-ttu-id="725b8-281">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-281">Type</span></span>| <span data-ttu-id="725b8-282">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="725b8-283">String</span><span class="sxs-lookup"><span data-stu-id="725b8-283">String</span></span>|<span data-ttu-id="725b8-284">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="725b8-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="725b8-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="725b8-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="725b8-286">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="725b8-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-287">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-287">Requirements</span></span>

|<span data-ttu-id="725b8-288">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-288">Requirement</span></span>| <span data-ttu-id="725b8-289">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-290">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-291">1.3</span><span class="sxs-lookup"><span data-stu-id="725b8-291">1.3</span></span>|
|[<span data-ttu-id="725b8-292">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-293">Restreinte</span><span class="sxs-lookup"><span data-stu-id="725b8-293">Restricted</span></span>|
|[<span data-ttu-id="725b8-294">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-295">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="725b8-296">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="725b8-296">Returns:</span></span>

<span data-ttu-id="725b8-297">Type : String</span><span class="sxs-lookup"><span data-stu-id="725b8-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="725b8-298">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="725b8-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="725b8-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="725b8-300">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="725b8-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="725b8-301">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="725b8-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-302">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-302">Parameters</span></span>

|<span data-ttu-id="725b8-303">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-303">Name</span></span>| <span data-ttu-id="725b8-304">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-304">Type</span></span>| <span data-ttu-id="725b8-305">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="725b8-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="725b8-306">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="725b8-307">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="725b8-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-308">Requirements</span></span>

|<span data-ttu-id="725b8-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-309">Requirement</span></span>| <span data-ttu-id="725b8-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-312">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-312">1.0</span></span>|
|[<span data-ttu-id="725b8-313">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-314">ReadItem</span></span>|
|[<span data-ttu-id="725b8-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-316">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="725b8-317">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="725b8-317">Returns:</span></span>

<span data-ttu-id="725b8-318">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="725b8-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="725b8-319">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="725b8-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="725b8-320">Date</span><span class="sxs-lookup"><span data-stu-id="725b8-320">Date</span></span></dd>

</dl>

---
---

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="725b8-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="725b8-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="725b8-322">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="725b8-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-323">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="725b8-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="725b8-324">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="725b8-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="725b8-p110">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="725b8-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="725b8-327">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="725b8-328">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="725b8-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-329">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-329">Parameters</span></span>

|<span data-ttu-id="725b8-330">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-330">Name</span></span>| <span data-ttu-id="725b8-331">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-331">Type</span></span>| <span data-ttu-id="725b8-332">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="725b8-333">String</span><span class="sxs-lookup"><span data-stu-id="725b8-333">String</span></span>|<span data-ttu-id="725b8-334">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="725b8-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-335">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-335">Requirements</span></span>

|<span data-ttu-id="725b8-336">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-336">Requirement</span></span>| <span data-ttu-id="725b8-337">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-338">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-339">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-339">1.0</span></span>|
|[<span data-ttu-id="725b8-340">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-341">ReadItem</span></span>|
|[<span data-ttu-id="725b8-342">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-343">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-344">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

####  <a name="displaymessageformitemid"></a><span data-ttu-id="725b8-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="725b8-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="725b8-346">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="725b8-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-347">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="725b8-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="725b8-348">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="725b8-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="725b8-349">Dans Outlook Web App, cette méthode ouvre le formulaire indiqué uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="725b8-350">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="725b8-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="725b8-p111">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="725b8-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-353">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-353">Parameters</span></span>

|<span data-ttu-id="725b8-354">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-354">Name</span></span>| <span data-ttu-id="725b8-355">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-355">Type</span></span>| <span data-ttu-id="725b8-356">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="725b8-357">String</span><span class="sxs-lookup"><span data-stu-id="725b8-357">String</span></span>|<span data-ttu-id="725b8-358">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="725b8-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-359">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-359">Requirements</span></span>

|<span data-ttu-id="725b8-360">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-360">Requirement</span></span>| <span data-ttu-id="725b8-361">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-362">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-363">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-363">1.0</span></span>|
|[<span data-ttu-id="725b8-364">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-365">ReadItem</span></span>|
|[<span data-ttu-id="725b8-366">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-367">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-368">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="725b8-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="725b8-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="725b8-370">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="725b8-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-371">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="725b8-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="725b8-p112">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="725b8-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="725b8-p113">Dans Outlook Web App et OWA pour les périphériques, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="725b8-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="725b8-p114">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="725b8-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="725b8-379">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="725b8-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-380">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-380">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-381">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="725b8-381">All parameters are optional.</span></span>

|<span data-ttu-id="725b8-382">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-382">Name</span></span>| <span data-ttu-id="725b8-383">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-383">Type</span></span>| <span data-ttu-id="725b8-384">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="725b8-385">Object</span><span class="sxs-lookup"><span data-stu-id="725b8-385">Object</span></span> | <span data-ttu-id="725b8-386">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="725b8-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="725b8-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="725b8-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="725b8-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="725b8-p116">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="725b8-393">Date</span><span class="sxs-lookup"><span data-stu-id="725b8-393">Date</span></span> | <span data-ttu-id="725b8-394">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="725b8-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="725b8-395">Date</span><span class="sxs-lookup"><span data-stu-id="725b8-395">Date</span></span> | <span data-ttu-id="725b8-396">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="725b8-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="725b8-397">Chaîne</span><span class="sxs-lookup"><span data-stu-id="725b8-397">String</span></span> | <span data-ttu-id="725b8-p117">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="725b8-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="725b8-p118">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="725b8-403">String</span><span class="sxs-lookup"><span data-stu-id="725b8-403">String</span></span> | <span data-ttu-id="725b8-p119">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="725b8-406">String</span><span class="sxs-lookup"><span data-stu-id="725b8-406">String</span></span> | <span data-ttu-id="725b8-p120">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="725b8-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="725b8-409">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-409">Requirements</span></span>

|<span data-ttu-id="725b8-410">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-410">Requirement</span></span>| <span data-ttu-id="725b8-411">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-412">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-413">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-413">1.0</span></span>|
|[<span data-ttu-id="725b8-414">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-414">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-415">ReadItem</span></span>|
|[<span data-ttu-id="725b8-416">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-416">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-417">Lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-418">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="725b8-419">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="725b8-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="725b8-420">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="725b8-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="725b8-421">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l'utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="725b8-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="725b8-422">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="725b8-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="725b8-423">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="725b8-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-424">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-424">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-425">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="725b8-425">All parameters are optional.</span></span>

|<span data-ttu-id="725b8-426">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-426">Name</span></span>| <span data-ttu-id="725b8-427">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-427">Type</span></span>| <span data-ttu-id="725b8-428">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="725b8-429">Object</span><span class="sxs-lookup"><span data-stu-id="725b8-429">Object</span></span> | <span data-ttu-id="725b8-430">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="725b8-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="725b8-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="725b8-432">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="725b8-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="725b8-433">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="725b8-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="725b8-435">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="725b8-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="725b8-436">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="725b8-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="725b8-438">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="725b8-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="725b8-439">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="725b8-440">String</span><span class="sxs-lookup"><span data-stu-id="725b8-440">String</span></span> | <span data-ttu-id="725b8-441">Chaîne contenant l'objet du message.</span><span class="sxs-lookup"><span data-stu-id="725b8-441">A string containing the subject of the message.</span></span> <span data-ttu-id="725b8-442">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="725b8-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="725b8-443">Chaîne</span><span class="sxs-lookup"><span data-stu-id="725b8-443">String</span></span> | <span data-ttu-id="725b8-444">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="725b8-444">The HTML body of the message.</span></span> <span data-ttu-id="725b8-445">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="725b8-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="725b8-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="725b8-447">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="725b8-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="725b8-448">Chaîne</span><span class="sxs-lookup"><span data-stu-id="725b8-448">String</span></span> | <span data-ttu-id="725b8-p127">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="725b8-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="725b8-451">String</span><span class="sxs-lookup"><span data-stu-id="725b8-451">String</span></span> | <span data-ttu-id="725b8-452">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="725b8-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="725b8-453">Chaîne</span><span class="sxs-lookup"><span data-stu-id="725b8-453">String</span></span> | <span data-ttu-id="725b8-p128">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="725b8-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="725b8-456">Booléen</span><span class="sxs-lookup"><span data-stu-id="725b8-456">Boolean</span></span> | <span data-ttu-id="725b8-p129">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="725b8-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="725b8-459">String</span><span class="sxs-lookup"><span data-stu-id="725b8-459">String</span></span> | <span data-ttu-id="725b8-460">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="725b8-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="725b8-461">ID d'élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="725b8-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="725b8-462">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="725b8-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="725b8-463">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-463">Requirements</span></span>

|<span data-ttu-id="725b8-464">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-464">Requirement</span></span>| <span data-ttu-id="725b8-465">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-466">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-467">1.6</span><span class="sxs-lookup"><span data-stu-id="725b8-467">1.6</span></span> |
|[<span data-ttu-id="725b8-468">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-469">ReadItem</span></span>|
|[<span data-ttu-id="725b8-470">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-471">Lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-472">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-472">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="725b8-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="725b8-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="725b8-474">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="725b8-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="725b8-p131">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="725b8-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-477">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="725b8-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="725b8-478">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="725b8-478">**REST Tokens**</span></span>

<span data-ttu-id="725b8-p132">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="725b8-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="725b8-482">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="725b8-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="725b8-483">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="725b8-483">**EWS Tokens**</span></span>

<span data-ttu-id="725b8-p133">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="725b8-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="725b8-486">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="725b8-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-487">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-487">Parameters</span></span>

|<span data-ttu-id="725b8-488">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-488">Name</span></span>| <span data-ttu-id="725b8-489">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-489">Type</span></span>| <span data-ttu-id="725b8-490">Attributs</span><span class="sxs-lookup"><span data-stu-id="725b8-490">Attributes</span></span>| <span data-ttu-id="725b8-491">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="725b8-492">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-492">Object</span></span> | <span data-ttu-id="725b8-493">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-493">&lt;optional&gt;</span></span> | <span data-ttu-id="725b8-494">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="725b8-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="725b8-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="725b8-495">Boolean</span></span> |  <span data-ttu-id="725b8-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-496">&lt;optional&gt;</span></span> | <span data-ttu-id="725b8-p134">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="725b8-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="725b8-499">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-499">Object</span></span> |  <span data-ttu-id="725b8-500">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-500">&lt;optional&gt;</span></span> | <span data-ttu-id="725b8-501">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="725b8-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="725b8-502">fonction</span><span class="sxs-lookup"><span data-stu-id="725b8-502">function</span></span>||<span data-ttu-id="725b8-p135">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="725b8-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-505">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-505">Requirements</span></span>

|<span data-ttu-id="725b8-506">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-506">Requirement</span></span>| <span data-ttu-id="725b8-507">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-508">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-509">1,5</span><span class="sxs-lookup"><span data-stu-id="725b8-509">1.5</span></span> |
|[<span data-ttu-id="725b8-510">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-511">ReadItem</span></span>|
|[<span data-ttu-id="725b8-512">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-513">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-514">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="725b8-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="725b8-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="725b8-516">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="725b8-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="725b8-p136">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="725b8-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="725b8-p137">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="725b8-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="725b8-522">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="725b8-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="725b8-p138">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="725b8-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-525">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-525">Parameters</span></span>

|<span data-ttu-id="725b8-526">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-526">Name</span></span>| <span data-ttu-id="725b8-527">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-527">Type</span></span>| <span data-ttu-id="725b8-528">Attributs</span><span class="sxs-lookup"><span data-stu-id="725b8-528">Attributes</span></span>| <span data-ttu-id="725b8-529">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="725b8-530">fonction</span><span class="sxs-lookup"><span data-stu-id="725b8-530">function</span></span>||<span data-ttu-id="725b8-p139">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="725b8-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="725b8-533">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-533">Object</span></span>| <span data-ttu-id="725b8-534">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-534">&lt;optional&gt;</span></span>|<span data-ttu-id="725b8-535">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="725b8-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-536">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-536">Requirements</span></span>

|<span data-ttu-id="725b8-537">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-537">Requirement</span></span>| <span data-ttu-id="725b8-538">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-539">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-540">1.3</span><span class="sxs-lookup"><span data-stu-id="725b8-540">1.3</span></span>|
|[<span data-ttu-id="725b8-541">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-542">ReadItem</span></span>|
|[<span data-ttu-id="725b8-543">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-544">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-545">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-545">Example</span></span>

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

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="725b8-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="725b8-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="725b8-547">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="725b8-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="725b8-548">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="725b8-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-549">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-549">Parameters</span></span>

|<span data-ttu-id="725b8-550">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-550">Name</span></span>| <span data-ttu-id="725b8-551">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-551">Type</span></span>| <span data-ttu-id="725b8-552">Attributs</span><span class="sxs-lookup"><span data-stu-id="725b8-552">Attributes</span></span>| <span data-ttu-id="725b8-553">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="725b8-554">function</span><span class="sxs-lookup"><span data-stu-id="725b8-554">function</span></span>||<span data-ttu-id="725b8-555">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="725b8-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="725b8-556">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="725b8-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="725b8-557">Object</span><span class="sxs-lookup"><span data-stu-id="725b8-557">Object</span></span>| <span data-ttu-id="725b8-558">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-558">&lt;optional&gt;</span></span>|<span data-ttu-id="725b8-559">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="725b8-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-560">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-560">Requirements</span></span>

|<span data-ttu-id="725b8-561">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-561">Requirement</span></span>| <span data-ttu-id="725b8-562">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-563">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-564">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-564">1.0</span></span>|
|[<span data-ttu-id="725b8-565">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-566">ReadItem</span></span>|
|[<span data-ttu-id="725b8-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-568">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-568">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-569">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-569">Example</span></span>

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

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="725b8-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="725b8-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="725b8-571">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="725b8-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-572">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="725b8-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="725b8-573">dans Outlook pour iOS ou Outlook pour Android ;</span><span class="sxs-lookup"><span data-stu-id="725b8-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="725b8-574">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="725b8-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="725b8-575">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="725b8-575">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="725b8-576">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="725b8-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="725b8-577">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="725b8-577">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="725b8-578">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="725b8-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="725b8-579">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="725b8-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="725b8-p141">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="725b8-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="725b8-582">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="725b8-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="725b8-583">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="725b8-583">Version differences</span></span>

<span data-ttu-id="725b8-584">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="725b8-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="725b8-p142">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="725b8-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-588">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-588">Parameters</span></span>

|<span data-ttu-id="725b8-589">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-589">Name</span></span>| <span data-ttu-id="725b8-590">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-590">Type</span></span>| <span data-ttu-id="725b8-591">Attributs</span><span class="sxs-lookup"><span data-stu-id="725b8-591">Attributes</span></span>| <span data-ttu-id="725b8-592">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="725b8-593">String</span><span class="sxs-lookup"><span data-stu-id="725b8-593">String</span></span>||<span data-ttu-id="725b8-594">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="725b8-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="725b8-595">function</span><span class="sxs-lookup"><span data-stu-id="725b8-595">function</span></span>||<span data-ttu-id="725b8-596">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="725b8-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="725b8-597">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="725b8-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="725b8-598">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="725b8-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="725b8-599">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-599">Object</span></span>| <span data-ttu-id="725b8-600">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-600">&lt;optional&gt;</span></span>|<span data-ttu-id="725b8-601">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="725b8-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-602">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-602">Requirements</span></span>

|<span data-ttu-id="725b8-603">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-603">Requirement</span></span>| <span data-ttu-id="725b8-604">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-605">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-606">1.0</span><span class="sxs-lookup"><span data-stu-id="725b8-606">1.0</span></span>|
|[<span data-ttu-id="725b8-607">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-607">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="725b8-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="725b8-609">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-609">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-610">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-610">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="725b8-611">Exemple</span><span class="sxs-lookup"><span data-stu-id="725b8-611">Example</span></span>

<span data-ttu-id="725b8-612">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="725b8-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="725b8-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="725b8-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="725b8-614">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="725b8-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="725b8-615">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="725b8-615">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="725b8-616">Paramètres</span><span class="sxs-lookup"><span data-stu-id="725b8-616">Parameters</span></span>

| <span data-ttu-id="725b8-617">Nom</span><span class="sxs-lookup"><span data-stu-id="725b8-617">Name</span></span> | <span data-ttu-id="725b8-618">Type</span><span class="sxs-lookup"><span data-stu-id="725b8-618">Type</span></span> | <span data-ttu-id="725b8-619">Attributs</span><span class="sxs-lookup"><span data-stu-id="725b8-619">Attributes</span></span> | <span data-ttu-id="725b8-620">Description</span><span class="sxs-lookup"><span data-stu-id="725b8-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="725b8-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="725b8-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="725b8-622">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="725b8-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="725b8-623">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-623">Object</span></span> | <span data-ttu-id="725b8-624">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-624">&lt;optional&gt;</span></span> | <span data-ttu-id="725b8-625">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="725b8-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="725b8-626">Objet</span><span class="sxs-lookup"><span data-stu-id="725b8-626">Object</span></span> | <span data-ttu-id="725b8-627">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-627">&lt;optional&gt;</span></span> | <span data-ttu-id="725b8-628">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="725b8-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="725b8-629">fonction</span><span class="sxs-lookup"><span data-stu-id="725b8-629">function</span></span>| <span data-ttu-id="725b8-630">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="725b8-630">&lt;optional&gt;</span></span>|<span data-ttu-id="725b8-631">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="725b8-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="725b8-632">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="725b8-632">Requirements</span></span>

|<span data-ttu-id="725b8-633">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="725b8-633">Requirement</span></span>| <span data-ttu-id="725b8-634">Valeur</span><span class="sxs-lookup"><span data-stu-id="725b8-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="725b8-635">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="725b8-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="725b8-636">1,5</span><span class="sxs-lookup"><span data-stu-id="725b8-636">1.5</span></span> |
|[<span data-ttu-id="725b8-637">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="725b8-637">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="725b8-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="725b8-638">ReadItem</span></span> |
|[<span data-ttu-id="725b8-639">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="725b8-639">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="725b8-640">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="725b8-640">Compose or Read</span></span>|
