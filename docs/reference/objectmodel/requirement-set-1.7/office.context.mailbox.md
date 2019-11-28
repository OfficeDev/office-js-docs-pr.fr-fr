---
title: Office. Context. Mailbox-ensemble de conditions requises 1,7
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: c310ad38bb9821955fb0571d3693ce39715376f4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629671"
---
# <a name="mailbox"></a><span data-ttu-id="4f859-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="4f859-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="4f859-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="4f859-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f859-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f859-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-105">Requirements</span></span>

|<span data-ttu-id="4f859-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-106">Requirement</span></span>| <span data-ttu-id="4f859-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-109">1.0</span></span>|
|[<span data-ttu-id="4f859-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4f859-111">Restricted</span></span>|
|[<span data-ttu-id="4f859-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4f859-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4f859-114">Members and methods</span></span>

| <span data-ttu-id="4f859-115">Membre</span><span class="sxs-lookup"><span data-stu-id="4f859-115">Member</span></span> | <span data-ttu-id="4f859-116">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4f859-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="4f859-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="4f859-118">Membre</span><span class="sxs-lookup"><span data-stu-id="4f859-118">Member</span></span> |
| [<span data-ttu-id="4f859-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="4f859-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="4f859-120">Membre</span><span class="sxs-lookup"><span data-stu-id="4f859-120">Member</span></span> |
| [<span data-ttu-id="4f859-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4f859-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="4f859-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-122">Method</span></span> |
| [<span data-ttu-id="4f859-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="4f859-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="4f859-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-124">Method</span></span> |
| [<span data-ttu-id="4f859-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4f859-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="4f859-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-126">Method</span></span> |
| [<span data-ttu-id="4f859-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="4f859-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="4f859-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-128">Method</span></span> |
| [<span data-ttu-id="4f859-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="4f859-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="4f859-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-130">Method</span></span> |
| [<span data-ttu-id="4f859-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4f859-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="4f859-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-132">Method</span></span> |
| [<span data-ttu-id="4f859-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="4f859-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="4f859-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-134">Method</span></span> |
| [<span data-ttu-id="4f859-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="4f859-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="4f859-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-136">Method</span></span> |
| [<span data-ttu-id="4f859-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="4f859-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="4f859-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-138">Method</span></span> |
| [<span data-ttu-id="4f859-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f859-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="4f859-140">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-140">Method</span></span> |
| [<span data-ttu-id="4f859-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f859-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="4f859-142">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-142">Method</span></span> |
| [<span data-ttu-id="4f859-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="4f859-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="4f859-144">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-144">Method</span></span> |
| [<span data-ttu-id="4f859-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="4f859-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="4f859-146">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-146">Method</span></span> |
| [<span data-ttu-id="4f859-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="4f859-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="4f859-148">Méthode</span><span class="sxs-lookup"><span data-stu-id="4f859-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4f859-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4f859-149">Namespaces</span></span>

<span data-ttu-id="4f859-150">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f859-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4f859-151">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f859-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4f859-152">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4f859-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4f859-153">Members</span><span class="sxs-lookup"><span data-stu-id="4f859-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4f859-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="4f859-154">ewsUrl: String</span></span>

<span data-ttu-id="4f859-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4f859-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-157">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4f859-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f859-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4f859-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4f859-160">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="4f859-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="4f859-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4f859-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4f859-163">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-163">Type</span></span>

*   <span data-ttu-id="4f859-164">String</span><span class="sxs-lookup"><span data-stu-id="4f859-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f859-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-165">Requirements</span></span>

|<span data-ttu-id="4f859-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-166">Requirement</span></span>| <span data-ttu-id="4f859-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-169">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-169">1.0</span></span>|
|[<span data-ttu-id="4f859-170">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-171">ReadItem</span></span>|
|[<span data-ttu-id="4f859-172">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-173">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="4f859-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="4f859-174">restUrl: String</span></span>

<span data-ttu-id="4f859-175">obtient l’URL du point de terminaison REST de ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="4f859-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="4f859-176">La valeur `restUrl` peut être utilisée pour que l’[API REST](/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4f859-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="4f859-177">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-177">Type</span></span>

*   <span data-ttu-id="4f859-178">String</span><span class="sxs-lookup"><span data-stu-id="4f859-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f859-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-179">Requirements</span></span>

|<span data-ttu-id="4f859-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-180">Requirement</span></span>| <span data-ttu-id="4f859-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-183">1,5</span><span class="sxs-lookup"><span data-stu-id="4f859-183">1.5</span></span> |
|[<span data-ttu-id="4f859-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-185">ReadItem</span></span>|
|[<span data-ttu-id="4f859-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4f859-188">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4f859-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="4f859-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4f859-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="4f859-190">ajoute un gestionnaire d’événements pour un événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4f859-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="4f859-191">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4f859-191">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-192">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-192">Parameters</span></span>

| <span data-ttu-id="4f859-193">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-193">Name</span></span> | <span data-ttu-id="4f859-194">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-194">Type</span></span> | <span data-ttu-id="4f859-195">Attributs</span><span class="sxs-lookup"><span data-stu-id="4f859-195">Attributes</span></span> | <span data-ttu-id="4f859-196">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4f859-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4f859-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4f859-198">Événement qui doit appeler le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4f859-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="4f859-199">Fonction</span><span class="sxs-lookup"><span data-stu-id="4f859-199">Function</span></span> || <span data-ttu-id="4f859-p104">Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="4f859-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="4f859-203">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-203">Object</span></span> | <span data-ttu-id="4f859-204">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-204">&lt;optional&gt;</span></span> | <span data-ttu-id="4f859-205">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4f859-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f859-206">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-206">Object</span></span> | <span data-ttu-id="4f859-207">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-207">&lt;optional&gt;</span></span> | <span data-ttu-id="4f859-208">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4f859-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4f859-209">fonction</span><span class="sxs-lookup"><span data-stu-id="4f859-209">function</span></span>| <span data-ttu-id="4f859-210">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-210">&lt;optional&gt;</span></span>|<span data-ttu-id="4f859-211">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f859-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-212">Requirements</span></span>

|<span data-ttu-id="4f859-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-213">Requirement</span></span>| <span data-ttu-id="4f859-214">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-216">1,5</span><span class="sxs-lookup"><span data-stu-id="4f859-216">1.5</span></span> |
|[<span data-ttu-id="4f859-217">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-218">ReadItem</span></span> |
|[<span data-ttu-id="4f859-219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-220">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-221">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-221">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="4f859-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4f859-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4f859-223">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="4f859-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-224">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4f859-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f859-p105">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="4f859-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-227">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-227">Parameters</span></span>

|<span data-ttu-id="4f859-228">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-228">Name</span></span>| <span data-ttu-id="4f859-229">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-229">Type</span></span>| <span data-ttu-id="4f859-230">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f859-231">String</span><span class="sxs-lookup"><span data-stu-id="4f859-231">String</span></span>|<span data-ttu-id="4f859-232">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="4f859-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="4f859-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4f859-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="4f859-234">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-235">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-235">Requirements</span></span>

|<span data-ttu-id="4f859-236">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-236">Requirement</span></span>| <span data-ttu-id="4f859-237">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-238">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-239">1.3</span><span class="sxs-lookup"><span data-stu-id="4f859-239">1.3</span></span>|
|[<span data-ttu-id="4f859-240">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-241">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4f859-241">Restricted</span></span>|
|[<span data-ttu-id="4f859-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f859-244">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4f859-244">Returns:</span></span>

<span data-ttu-id="4f859-245">Type : String</span><span class="sxs-lookup"><span data-stu-id="4f859-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4f859-246">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="4f859-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="4f859-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="4f859-248">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="4f859-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4f859-p106">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4f859-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4f859-p107">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="4f859-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-254">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4f859-254">Parameters</span></span>

|<span data-ttu-id="4f859-255">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-255">Name</span></span>| <span data-ttu-id="4f859-256">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-256">Type</span></span>| <span data-ttu-id="4f859-257">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4f859-258">Date</span><span class="sxs-lookup"><span data-stu-id="4f859-258">Date</span></span>|<span data-ttu-id="4f859-259">Objet Date</span><span class="sxs-lookup"><span data-stu-id="4f859-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-260">Requirements</span></span>

|<span data-ttu-id="4f859-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-261">Requirement</span></span>| <span data-ttu-id="4f859-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-264">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-264">1.0</span></span>|
|[<span data-ttu-id="4f859-265">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-266">ReadItem</span></span>|
|[<span data-ttu-id="4f859-267">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-268">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f859-269">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4f859-269">Returns:</span></span>

<span data-ttu-id="4f859-270">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="4f859-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="4f859-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4f859-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4f859-272">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="4f859-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-273">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4f859-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f859-p108">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="4f859-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-276">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-276">Parameters</span></span>

|<span data-ttu-id="4f859-277">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-277">Name</span></span>| <span data-ttu-id="4f859-278">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-278">Type</span></span>| <span data-ttu-id="4f859-279">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f859-280">String</span><span class="sxs-lookup"><span data-stu-id="4f859-280">String</span></span>|<span data-ttu-id="4f859-281">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="4f859-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="4f859-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4f859-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="4f859-283">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="4f859-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-284">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-284">Requirements</span></span>

|<span data-ttu-id="4f859-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-285">Requirement</span></span>| <span data-ttu-id="4f859-286">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-287">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-288">1.3</span><span class="sxs-lookup"><span data-stu-id="4f859-288">1.3</span></span>|
|[<span data-ttu-id="4f859-289">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-290">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4f859-290">Restricted</span></span>|
|[<span data-ttu-id="4f859-291">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-292">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f859-293">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4f859-293">Returns:</span></span>

<span data-ttu-id="4f859-294">Type : String</span><span class="sxs-lookup"><span data-stu-id="4f859-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4f859-295">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4f859-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4f859-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4f859-297">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="4f859-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4f859-298">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="4f859-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-299">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-299">Parameters</span></span>

|<span data-ttu-id="4f859-300">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-300">Name</span></span>| <span data-ttu-id="4f859-301">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-301">Type</span></span>| <span data-ttu-id="4f859-302">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4f859-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4f859-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="4f859-304">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="4f859-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-305">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-305">Requirements</span></span>

|<span data-ttu-id="4f859-306">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-306">Requirement</span></span>| <span data-ttu-id="4f859-307">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-308">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-309">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-309">1.0</span></span>|
|[<span data-ttu-id="4f859-310">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-311">ReadItem</span></span>|
|[<span data-ttu-id="4f859-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-313">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4f859-314">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4f859-314">Returns:</span></span>

<span data-ttu-id="4f859-315">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="4f859-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="4f859-316">Type : Date</span><span class="sxs-lookup"><span data-stu-id="4f859-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="4f859-317">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="4f859-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4f859-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4f859-319">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4f859-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-320">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4f859-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f859-321">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4f859-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4f859-p109">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="4f859-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4f859-324">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4f859-325">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4f859-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-326">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-326">Parameters</span></span>

|<span data-ttu-id="4f859-327">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-327">Name</span></span>| <span data-ttu-id="4f859-328">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-328">Type</span></span>| <span data-ttu-id="4f859-329">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f859-330">String</span><span class="sxs-lookup"><span data-stu-id="4f859-330">String</span></span>|<span data-ttu-id="4f859-331">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4f859-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-332">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-332">Requirements</span></span>

|<span data-ttu-id="4f859-333">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-333">Requirement</span></span>| <span data-ttu-id="4f859-334">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-335">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-336">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-336">1.0</span></span>|
|[<span data-ttu-id="4f859-337">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-338">ReadItem</span></span>|
|[<span data-ttu-id="4f859-339">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-340">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-341">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="4f859-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4f859-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4f859-343">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="4f859-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-344">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4f859-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f859-345">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4f859-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4f859-346">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4f859-347">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4f859-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4f859-p110">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4f859-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-350">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4f859-350">Parameters</span></span>

|<span data-ttu-id="4f859-351">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-351">Name</span></span>| <span data-ttu-id="4f859-352">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-352">Type</span></span>| <span data-ttu-id="4f859-353">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4f859-354">String</span><span class="sxs-lookup"><span data-stu-id="4f859-354">String</span></span>|<span data-ttu-id="4f859-355">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="4f859-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-356">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-356">Requirements</span></span>

|<span data-ttu-id="4f859-357">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-357">Requirement</span></span>| <span data-ttu-id="4f859-358">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-359">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-360">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-360">1.0</span></span>|
|[<span data-ttu-id="4f859-361">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-362">ReadItem</span></span>|
|[<span data-ttu-id="4f859-363">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-364">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-365">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4f859-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4f859-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4f859-367">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="4f859-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-368">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4f859-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="4f859-p111">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="4f859-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4f859-p112">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="4f859-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4f859-p113">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="4f859-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4f859-376">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="4f859-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-377">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4f859-377">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-378">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="4f859-378">All parameters are optional.</span></span>

|<span data-ttu-id="4f859-379">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-379">Name</span></span>| <span data-ttu-id="4f859-380">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-380">Type</span></span>| <span data-ttu-id="4f859-381">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4f859-382">Object</span><span class="sxs-lookup"><span data-stu-id="4f859-382">Object</span></span> | <span data-ttu-id="4f859-383">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4f859-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4f859-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f859-p114">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4f859-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f859-p115">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4f859-390">Date</span><span class="sxs-lookup"><span data-stu-id="4f859-390">Date</span></span> | <span data-ttu-id="4f859-391">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4f859-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4f859-392">Date</span><span class="sxs-lookup"><span data-stu-id="4f859-392">Date</span></span> | <span data-ttu-id="4f859-393">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4f859-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4f859-394">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4f859-394">String</span></span> | <span data-ttu-id="4f859-p116">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4f859-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4f859-p117">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4f859-400">String</span><span class="sxs-lookup"><span data-stu-id="4f859-400">String</span></span> | <span data-ttu-id="4f859-p118">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4f859-403">String</span><span class="sxs-lookup"><span data-stu-id="4f859-403">String</span></span> | <span data-ttu-id="4f859-p119">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4f859-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4f859-406">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-406">Requirements</span></span>

|<span data-ttu-id="4f859-407">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-407">Requirement</span></span>| <span data-ttu-id="4f859-408">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-409">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-410">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-410">1.0</span></span>|
|[<span data-ttu-id="4f859-411">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-412">ReadItem</span></span>|
|[<span data-ttu-id="4f859-413">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-414">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-415">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-415">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="4f859-416">displayNewMessageForm (paramètres)</span><span class="sxs-lookup"><span data-stu-id="4f859-416">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="4f859-417">Affiche un formulaire permettant de créer un message.</span><span class="sxs-lookup"><span data-stu-id="4f859-417">Displays a form for creating a new message.</span></span>

<span data-ttu-id="4f859-418">La `displayNewMessageForm` méthode ouvre un formulaire qui permet à l’utilisateur de créer un message.</span><span class="sxs-lookup"><span data-stu-id="4f859-418">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="4f859-419">Si les paramètres sont spécifiés, les champs du formulaire de message sont automatiquement renseignés avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="4f859-419">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4f859-420">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="4f859-420">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-421">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4f859-421">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-422">Tous les paramètres sont facultatifs.</span><span class="sxs-lookup"><span data-stu-id="4f859-422">All parameters are optional.</span></span>

|<span data-ttu-id="4f859-423">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-423">Name</span></span>| <span data-ttu-id="4f859-424">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-424">Type</span></span>| <span data-ttu-id="4f859-425">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-425">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4f859-426">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-426">Object</span></span> | <span data-ttu-id="4f859-427">Dictionnaire de paramètres décrivant le nouveau message.</span><span class="sxs-lookup"><span data-stu-id="4f859-427">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="4f859-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f859-429">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne à.</span><span class="sxs-lookup"><span data-stu-id="4f859-429">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="4f859-430">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-430">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="4f859-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f859-432">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CC.</span><span class="sxs-lookup"><span data-stu-id="4f859-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="4f859-433">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="4f859-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="4f859-435">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un `EmailAddressDetails` objet pour chacun des destinataires de la ligne CCI.</span><span class="sxs-lookup"><span data-stu-id="4f859-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="4f859-436">Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4f859-437">String</span><span class="sxs-lookup"><span data-stu-id="4f859-437">String</span></span> | <span data-ttu-id="4f859-438">Chaîne contenant l’objet du message.</span><span class="sxs-lookup"><span data-stu-id="4f859-438">A string containing the subject of the message.</span></span> <span data-ttu-id="4f859-439">La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4f859-439">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="4f859-440">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4f859-440">String</span></span> | <span data-ttu-id="4f859-441">Corps HTML du message.</span><span class="sxs-lookup"><span data-stu-id="4f859-441">The HTML body of the message.</span></span> <span data-ttu-id="4f859-442">La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4f859-442">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="4f859-443">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-443">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="4f859-444">Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-444">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="4f859-445">String</span><span class="sxs-lookup"><span data-stu-id="4f859-445">String</span></span> | <span data-ttu-id="4f859-p126">Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="4f859-448">String</span><span class="sxs-lookup"><span data-stu-id="4f859-448">String</span></span> | <span data-ttu-id="4f859-449">Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.</span><span class="sxs-lookup"><span data-stu-id="4f859-449">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="4f859-450">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4f859-450">String</span></span> | <span data-ttu-id="4f859-p127">Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.</span><span class="sxs-lookup"><span data-stu-id="4f859-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="4f859-453">Booléen</span><span class="sxs-lookup"><span data-stu-id="4f859-453">Boolean</span></span> | <span data-ttu-id="4f859-p128">Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.</span><span class="sxs-lookup"><span data-stu-id="4f859-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="4f859-456">String</span><span class="sxs-lookup"><span data-stu-id="4f859-456">String</span></span> | <span data-ttu-id="4f859-457">Utilisé uniquement si `type` est défini sur `item`.</span><span class="sxs-lookup"><span data-stu-id="4f859-457">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="4f859-458">ID d’élément EWS du message électronique existant que vous souhaitez joindre au nouveau message.</span><span class="sxs-lookup"><span data-stu-id="4f859-458">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="4f859-459">Il s’agit d’une chaîne comportant un maximum de 100 caractères.</span><span class="sxs-lookup"><span data-stu-id="4f859-459">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="4f859-460">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-460">Requirements</span></span>

|<span data-ttu-id="4f859-461">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-461">Requirement</span></span>| <span data-ttu-id="4f859-462">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-463">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-464">1.6</span><span class="sxs-lookup"><span data-stu-id="4f859-464">1.6</span></span> |
|[<span data-ttu-id="4f859-465">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-465">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-466">ReadItem</span></span>|
|[<span data-ttu-id="4f859-467">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-467">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-468">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-468">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-469">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-469">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="4f859-470">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="4f859-470">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="4f859-471">Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4f859-471">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="4f859-p130">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="4f859-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-474">Les compléments devraient, dans la mesure du possible, utiliser les API REST à la place des services web Exchange.</span><span class="sxs-lookup"><span data-stu-id="4f859-474">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="4f859-475">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="4f859-475">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="4f859-476">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-476">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="4f859-477">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="4f859-477">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="4f859-478">**Jetons REST**</span><span class="sxs-lookup"><span data-stu-id="4f859-478">**REST Tokens**</span></span>

<span data-ttu-id="4f859-p132">Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services web Exchange. Le jeton peut uniquement accéder en lecture seule à l’élément actif et à ses pièces jointes, sauf si l’autorisation [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.</span><span class="sxs-lookup"><span data-stu-id="4f859-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="4f859-482">Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.</span><span class="sxs-lookup"><span data-stu-id="4f859-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="4f859-483">**Jetons EWS**</span><span class="sxs-lookup"><span data-stu-id="4f859-483">**EWS Tokens**</span></span>

<span data-ttu-id="4f859-p133">Quand un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton peut uniquement accéder à l’élément actif.</span><span class="sxs-lookup"><span data-stu-id="4f859-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="4f859-486">Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.</span><span class="sxs-lookup"><span data-stu-id="4f859-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="4f859-487">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="4f859-487">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="4f859-488">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-488">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="4f859-489">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4f859-489">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-490">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-490">Parameters</span></span>

|<span data-ttu-id="4f859-491">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-491">Name</span></span>| <span data-ttu-id="4f859-492">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-492">Type</span></span>| <span data-ttu-id="4f859-493">Attributs</span><span class="sxs-lookup"><span data-stu-id="4f859-493">Attributes</span></span>| <span data-ttu-id="4f859-494">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-494">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="4f859-495">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-495">Object</span></span> | <span data-ttu-id="4f859-496">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-496">&lt;optional&gt;</span></span> | <span data-ttu-id="4f859-497">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4f859-497">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="4f859-498">Boolean</span><span class="sxs-lookup"><span data-stu-id="4f859-498">Boolean</span></span> |  <span data-ttu-id="4f859-499">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-499">&lt;optional&gt;</span></span> | <span data-ttu-id="4f859-p135">Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services web Exchange. La valeur par défaut est `false`.</span><span class="sxs-lookup"><span data-stu-id="4f859-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f859-502">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-502">Object</span></span> |  <span data-ttu-id="4f859-503">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-503">&lt;optional&gt;</span></span> | <span data-ttu-id="4f859-504">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4f859-504">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="4f859-505">fonction</span><span class="sxs-lookup"><span data-stu-id="4f859-505">function</span></span>||<span data-ttu-id="4f859-506">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f859-506">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f859-507">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f859-507">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f859-508">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4f859-508">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f859-509">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4f859-509">Errors</span></span>

|<span data-ttu-id="4f859-510">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4f859-510">Error code</span></span>|<span data-ttu-id="4f859-511">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-511">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f859-512">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4f859-512">The request has failed.</span></span> <span data-ttu-id="4f859-513">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4f859-513">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f859-514">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4f859-514">The Exchange server returned an error.</span></span> <span data-ttu-id="4f859-515">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4f859-515">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f859-516">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4f859-516">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f859-517">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="4f859-517">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-518">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-518">Requirements</span></span>

|<span data-ttu-id="4f859-519">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-519">Requirement</span></span>| <span data-ttu-id="4f859-520">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-521">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-522">1,5</span><span class="sxs-lookup"><span data-stu-id="4f859-522">1.5</span></span> |
|[<span data-ttu-id="4f859-523">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-523">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-524">ReadItem</span></span>|
|[<span data-ttu-id="4f859-525">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-525">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-526">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-526">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-527">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-527">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4f859-528">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f859-528">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4f859-529">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="4f859-529">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4f859-p139">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="4f859-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4f859-532">Vous pouvez passer à la fois le jeton et un identifiant de pièce jointe ou un identifiant d'élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="4f859-532">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="4f859-533">Le système tiers utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-533">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="4f859-534">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4f859-534">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4f859-535">L’appel de la méthode `getCallbackTokenAsync` en mode lecture nécessite un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="4f859-535">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="4f859-536">Pour appeler `getCallbackTokenAsync` en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-536">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="4f859-537">La méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) nécessite un niveau d’autorisation minimal de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="4f859-537">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-538">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-538">Parameters</span></span>

|<span data-ttu-id="4f859-539">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-539">Name</span></span>| <span data-ttu-id="4f859-540">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-540">Type</span></span>| <span data-ttu-id="4f859-541">Attributs</span><span class="sxs-lookup"><span data-stu-id="4f859-541">Attributes</span></span>| <span data-ttu-id="4f859-542">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-542">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4f859-543">function</span><span class="sxs-lookup"><span data-stu-id="4f859-543">function</span></span>||<span data-ttu-id="4f859-544">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f859-544">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f859-545">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f859-545">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f859-546">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4f859-546">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4f859-547">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-547">Object</span></span>| <span data-ttu-id="4f859-548">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-548">&lt;optional&gt;</span></span>|<span data-ttu-id="4f859-549">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4f859-549">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f859-550">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4f859-550">Errors</span></span>

|<span data-ttu-id="4f859-551">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4f859-551">Error code</span></span>|<span data-ttu-id="4f859-552">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-552">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f859-553">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4f859-553">The request has failed.</span></span> <span data-ttu-id="4f859-554">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4f859-554">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f859-555">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4f859-555">The Exchange server returned an error.</span></span> <span data-ttu-id="4f859-556">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4f859-556">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f859-557">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4f859-557">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f859-558">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="4f859-558">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-559">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-559">Requirements</span></span>

|<span data-ttu-id="4f859-560">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-560">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="4f859-561">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-562">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-562">1.0</span></span> | <span data-ttu-id="4f859-563">1.3</span><span class="sxs-lookup"><span data-stu-id="4f859-563">1.3</span></span> |
|[<span data-ttu-id="4f859-564">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-565">ReadItem</span></span> | <span data-ttu-id="4f859-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-566">ReadItem</span></span> |
|[<span data-ttu-id="4f859-567">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-568">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-568">Read</span></span> | <span data-ttu-id="4f859-569">Composition</span><span class="sxs-lookup"><span data-stu-id="4f859-569">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="4f859-570">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-570">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4f859-571">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f859-571">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4f859-572">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="4f859-572">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4f859-573">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="4f859-573">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-574">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4f859-574">Parameters</span></span>

|<span data-ttu-id="4f859-575">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-575">Name</span></span>| <span data-ttu-id="4f859-576">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-576">Type</span></span>| <span data-ttu-id="4f859-577">Attributs</span><span class="sxs-lookup"><span data-stu-id="4f859-577">Attributes</span></span>| <span data-ttu-id="4f859-578">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-578">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4f859-579">function</span><span class="sxs-lookup"><span data-stu-id="4f859-579">function</span></span>||<span data-ttu-id="4f859-580">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f859-580">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f859-581">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f859-581">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="4f859-582">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="4f859-582">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="4f859-583">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-583">Object</span></span>| <span data-ttu-id="4f859-584">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-584">&lt;optional&gt;</span></span>|<span data-ttu-id="4f859-585">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4f859-585">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="4f859-586">Erreurs</span><span class="sxs-lookup"><span data-stu-id="4f859-586">Errors</span></span>

|<span data-ttu-id="4f859-587">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="4f859-587">Error code</span></span>|<span data-ttu-id="4f859-588">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-588">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="4f859-589">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="4f859-589">The request has failed.</span></span> <span data-ttu-id="4f859-590">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4f859-590">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="4f859-591">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="4f859-591">The Exchange server returned an error.</span></span> <span data-ttu-id="4f859-592">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="4f859-592">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="4f859-593">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="4f859-593">The user is no longer connected to the network.</span></span> <span data-ttu-id="4f859-594">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="4f859-594">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-595">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-595">Requirements</span></span>

|<span data-ttu-id="4f859-596">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-596">Requirement</span></span>| <span data-ttu-id="4f859-597">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-597">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-598">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-598">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-599">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-599">1.0</span></span>|
|[<span data-ttu-id="4f859-600">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-600">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-601">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-601">ReadItem</span></span>|
|[<span data-ttu-id="4f859-602">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-602">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-603">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-603">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-604">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-604">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4f859-605">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4f859-605">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4f859-606">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4f859-606">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-607">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="4f859-607">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4f859-608">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="4f859-608">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="4f859-609">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="4f859-609">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4f859-610">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4f859-610">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4f859-611">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="4f859-611">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4f859-612">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="4f859-612">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4f859-613">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="4f859-613">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4f859-614">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="4f859-614">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4f859-p149">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="4f859-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4f859-617">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="4f859-617">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4f859-618">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="4f859-618">Version differences</span></span>

<span data-ttu-id="4f859-619">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="4f859-619">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4f859-p150">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="4f859-p150">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-623">Parameters</span><span class="sxs-lookup"><span data-stu-id="4f859-623">Parameters</span></span>

|<span data-ttu-id="4f859-624">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-624">Name</span></span>| <span data-ttu-id="4f859-625">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-625">Type</span></span>| <span data-ttu-id="4f859-626">Attributs</span><span class="sxs-lookup"><span data-stu-id="4f859-626">Attributes</span></span>| <span data-ttu-id="4f859-627">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-627">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4f859-628">String</span><span class="sxs-lookup"><span data-stu-id="4f859-628">String</span></span>||<span data-ttu-id="4f859-629">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="4f859-629">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4f859-630">function</span><span class="sxs-lookup"><span data-stu-id="4f859-630">function</span></span>||<span data-ttu-id="4f859-631">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f859-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4f859-632">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4f859-632">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4f859-633">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="4f859-633">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="4f859-634">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-634">Object</span></span>| <span data-ttu-id="4f859-635">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-635">&lt;optional&gt;</span></span>|<span data-ttu-id="4f859-636">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4f859-636">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-637">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-637">Requirements</span></span>

|<span data-ttu-id="4f859-638">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-638">Requirement</span></span>| <span data-ttu-id="4f859-639">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-639">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-640">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-640">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-641">1.0</span><span class="sxs-lookup"><span data-stu-id="4f859-641">1.0</span></span>|
|[<span data-ttu-id="4f859-642">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-642">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-643">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4f859-643">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4f859-644">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-644">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-645">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-645">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4f859-646">Exemple</span><span class="sxs-lookup"><span data-stu-id="4f859-646">Example</span></span>

<span data-ttu-id="4f859-647">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4f859-647">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="4f859-648">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="4f859-648">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="4f859-649">Supprime les gestionnaires d’événements pour un type d’événement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4f859-649">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="4f859-650">Actuellement, seul le type d’événement `Office.EventType.ItemChanged` est pris en charge.</span><span class="sxs-lookup"><span data-stu-id="4f859-650">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4f859-651">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4f859-651">Parameters</span></span>

| <span data-ttu-id="4f859-652">Nom</span><span class="sxs-lookup"><span data-stu-id="4f859-652">Name</span></span> | <span data-ttu-id="4f859-653">Type</span><span class="sxs-lookup"><span data-stu-id="4f859-653">Type</span></span> | <span data-ttu-id="4f859-654">Attributs</span><span class="sxs-lookup"><span data-stu-id="4f859-654">Attributes</span></span> | <span data-ttu-id="4f859-655">Description</span><span class="sxs-lookup"><span data-stu-id="4f859-655">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="4f859-656">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="4f859-656">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="4f859-657">Événement qui doit révoquer le gestionnaire.</span><span class="sxs-lookup"><span data-stu-id="4f859-657">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="4f859-658">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-658">Object</span></span> | <span data-ttu-id="4f859-659">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-659">&lt;optional&gt;</span></span> | <span data-ttu-id="4f859-660">Littéral d’objet contenant une ou plusieurs des propriétés suivantes.</span><span class="sxs-lookup"><span data-stu-id="4f859-660">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="4f859-661">Objet</span><span class="sxs-lookup"><span data-stu-id="4f859-661">Object</span></span> | <span data-ttu-id="4f859-662">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-662">&lt;optional&gt;</span></span> | <span data-ttu-id="4f859-663">Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4f859-663">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="4f859-664">fonction</span><span class="sxs-lookup"><span data-stu-id="4f859-664">function</span></span>| <span data-ttu-id="4f859-665">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4f859-665">&lt;optional&gt;</span></span>|<span data-ttu-id="4f859-666">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4f859-666">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f859-667">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f859-667">Requirements</span></span>

|<span data-ttu-id="4f859-668">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f859-668">Requirement</span></span>| <span data-ttu-id="4f859-669">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f859-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f859-670">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f859-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4f859-671">1,5</span><span class="sxs-lookup"><span data-stu-id="4f859-671">1.5</span></span> |
|[<span data-ttu-id="4f859-672">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4f859-672">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4f859-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4f859-673">ReadItem</span></span> |
|[<span data-ttu-id="4f859-674">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f859-674">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4f859-675">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f859-675">Compose or Read</span></span>|
