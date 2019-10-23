---
title: Office. Context. Mailbox-ensemble de conditions requises 1,3
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: f1896803c38abd03f63b0a9ae689d91eeb5540de
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627019"
---
# <a name="mailbox"></a><span data-ttu-id="a6687-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="a6687-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="a6687-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="a6687-104">Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6687-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6687-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-105">Requirements</span></span>

|<span data-ttu-id="a6687-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-106">Requirement</span></span>| <span data-ttu-id="a6687-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-109">1.0</span></span>|
|[<span data-ttu-id="a6687-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a6687-111">Restricted</span></span>|
|[<span data-ttu-id="a6687-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a6687-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="a6687-114">Members and methods</span></span>

| <span data-ttu-id="a6687-115">Membre</span><span class="sxs-lookup"><span data-stu-id="a6687-115">Member</span></span> | <span data-ttu-id="a6687-116">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a6687-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="a6687-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="a6687-118">Membre</span><span class="sxs-lookup"><span data-stu-id="a6687-118">Member</span></span> |
| [<span data-ttu-id="a6687-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="a6687-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="a6687-120">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-120">Method</span></span> |
| [<span data-ttu-id="a6687-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="a6687-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="a6687-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-122">Method</span></span> |
| [<span data-ttu-id="a6687-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="a6687-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="a6687-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-124">Method</span></span> |
| [<span data-ttu-id="a6687-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="a6687-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="a6687-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-126">Method</span></span> |
| [<span data-ttu-id="a6687-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="a6687-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="a6687-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-128">Method</span></span> |
| [<span data-ttu-id="a6687-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="a6687-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="a6687-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-130">Method</span></span> |
| [<span data-ttu-id="a6687-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="a6687-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="a6687-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-132">Method</span></span> |
| [<span data-ttu-id="a6687-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a6687-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="a6687-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-134">Method</span></span> |
| [<span data-ttu-id="a6687-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="a6687-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="a6687-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-136">Method</span></span> |
| [<span data-ttu-id="a6687-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="a6687-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="a6687-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="a6687-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a6687-139">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="a6687-139">Namespaces</span></span>

<span data-ttu-id="a6687-140">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6687-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="a6687-141">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6687-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="a6687-142">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a6687-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="a6687-143">Members</span><span class="sxs-lookup"><span data-stu-id="a6687-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="a6687-144">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="a6687-144">ewsUrl: String</span></span>

<span data-ttu-id="a6687-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="a6687-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-147">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a6687-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a6687-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="a6687-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="a6687-150">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="a6687-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="a6687-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="a6687-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="a6687-153">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-153">Type</span></span>

*   <span data-ttu-id="a6687-154">String</span><span class="sxs-lookup"><span data-stu-id="a6687-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6687-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-155">Requirements</span></span>

|<span data-ttu-id="a6687-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-156">Requirement</span></span>| <span data-ttu-id="a6687-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-159">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-159">1.0</span></span>|
|[<span data-ttu-id="a6687-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-161">ReadItem</span></span>|
|[<span data-ttu-id="a6687-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="a6687-164">Méthodes</span><span class="sxs-lookup"><span data-stu-id="a6687-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="a6687-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="a6687-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="a6687-166">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="a6687-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-167">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a6687-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a6687-p104">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="a6687-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-170">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-170">Parameters</span></span>

|<span data-ttu-id="a6687-171">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-171">Name</span></span>| <span data-ttu-id="a6687-172">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-172">Type</span></span>| <span data-ttu-id="a6687-173">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a6687-174">String</span><span class="sxs-lookup"><span data-stu-id="a6687-174">String</span></span>|<span data-ttu-id="a6687-175">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="a6687-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="a6687-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="a6687-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="a6687-177">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="a6687-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-178">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-178">Requirements</span></span>

|<span data-ttu-id="a6687-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-179">Requirement</span></span>| <span data-ttu-id="a6687-180">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-181">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-182">1.3</span><span class="sxs-lookup"><span data-stu-id="a6687-182">1.3</span></span>|
|[<span data-ttu-id="a6687-183">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-184">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a6687-184">Restricted</span></span>|
|[<span data-ttu-id="a6687-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-186">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6687-187">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a6687-187">Returns:</span></span>

<span data-ttu-id="a6687-188">Type : String</span><span class="sxs-lookup"><span data-stu-id="a6687-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a6687-189">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="a6687-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="a6687-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="a6687-191">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="a6687-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="a6687-p105">Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a6687-p105">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="a6687-p106">Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="a6687-p106">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-197">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-197">Parameters</span></span>

|<span data-ttu-id="a6687-198">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-198">Name</span></span>| <span data-ttu-id="a6687-199">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-199">Type</span></span>| <span data-ttu-id="a6687-200">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="a6687-201">Date</span><span class="sxs-lookup"><span data-stu-id="a6687-201">Date</span></span>|<span data-ttu-id="a6687-202">Objet Date</span><span class="sxs-lookup"><span data-stu-id="a6687-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-203">Requirements</span></span>

|<span data-ttu-id="a6687-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-204">Requirement</span></span>| <span data-ttu-id="a6687-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-207">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-207">1.0</span></span>|
|[<span data-ttu-id="a6687-208">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-209">ReadItem</span></span>|
|[<span data-ttu-id="a6687-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-211">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6687-212">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a6687-212">Returns:</span></span>

<span data-ttu-id="a6687-213">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="a6687-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="a6687-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="a6687-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="a6687-215">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="a6687-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-216">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a6687-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a6687-p107">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="a6687-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-219">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-219">Parameters</span></span>

|<span data-ttu-id="a6687-220">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-220">Name</span></span>| <span data-ttu-id="a6687-221">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-221">Type</span></span>| <span data-ttu-id="a6687-222">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a6687-223">String</span><span class="sxs-lookup"><span data-stu-id="a6687-223">String</span></span>|<span data-ttu-id="a6687-224">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="a6687-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="a6687-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="a6687-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="a6687-226">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="a6687-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-227">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-227">Requirements</span></span>

|<span data-ttu-id="a6687-228">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-228">Requirement</span></span>| <span data-ttu-id="a6687-229">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-230">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-231">1.3</span><span class="sxs-lookup"><span data-stu-id="a6687-231">1.3</span></span>|
|[<span data-ttu-id="a6687-232">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-233">Restreinte</span><span class="sxs-lookup"><span data-stu-id="a6687-233">Restricted</span></span>|
|[<span data-ttu-id="a6687-234">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-235">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6687-236">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a6687-236">Returns:</span></span>

<span data-ttu-id="a6687-237">Type : String</span><span class="sxs-lookup"><span data-stu-id="a6687-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="a6687-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="a6687-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="a6687-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="a6687-240">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="a6687-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="a6687-241">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="a6687-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-242">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-242">Parameters</span></span>

|<span data-ttu-id="a6687-243">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-243">Name</span></span>| <span data-ttu-id="a6687-244">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-244">Type</span></span>| <span data-ttu-id="a6687-245">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="a6687-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="a6687-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="a6687-247">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="a6687-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-248">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-248">Requirements</span></span>

|<span data-ttu-id="a6687-249">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-249">Requirement</span></span>| <span data-ttu-id="a6687-250">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-251">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-252">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-252">1.0</span></span>|
|[<span data-ttu-id="a6687-253">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-254">ReadItem</span></span>|
|[<span data-ttu-id="a6687-255">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-256">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="a6687-257">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="a6687-257">Returns:</span></span>

<span data-ttu-id="a6687-258">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="a6687-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="a6687-259">Type : Date</span><span class="sxs-lookup"><span data-stu-id="a6687-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="a6687-260">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="a6687-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="a6687-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="a6687-262">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="a6687-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-263">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a6687-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a6687-264">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="a6687-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="a6687-p108">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="a6687-p108">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="a6687-267">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="a6687-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="a6687-268">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="a6687-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-269">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-269">Parameters</span></span>

|<span data-ttu-id="a6687-270">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-270">Name</span></span>| <span data-ttu-id="a6687-271">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-271">Type</span></span>| <span data-ttu-id="a6687-272">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a6687-273">String</span><span class="sxs-lookup"><span data-stu-id="a6687-273">String</span></span>|<span data-ttu-id="a6687-274">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="a6687-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-275">Requirements</span></span>

|<span data-ttu-id="a6687-276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-276">Requirement</span></span>| <span data-ttu-id="a6687-277">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-279">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-279">1.0</span></span>|
|[<span data-ttu-id="a6687-280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-281">ReadItem</span></span>|
|[<span data-ttu-id="a6687-282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-283">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6687-284">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="a6687-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="a6687-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="a6687-286">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="a6687-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-287">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a6687-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a6687-288">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="a6687-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="a6687-289">Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="a6687-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="a6687-290">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="a6687-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="a6687-p109">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a6687-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-293">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-293">Parameters</span></span>

|<span data-ttu-id="a6687-294">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-294">Name</span></span>| <span data-ttu-id="a6687-295">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-295">Type</span></span>| <span data-ttu-id="a6687-296">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="a6687-297">String</span><span class="sxs-lookup"><span data-stu-id="a6687-297">String</span></span>|<span data-ttu-id="a6687-298">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="a6687-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-299">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-299">Requirements</span></span>

|<span data-ttu-id="a6687-300">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-300">Requirement</span></span>| <span data-ttu-id="a6687-301">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-302">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-303">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-303">1.0</span></span>|
|[<span data-ttu-id="a6687-304">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-305">ReadItem</span></span>|
|[<span data-ttu-id="a6687-306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-307">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6687-308">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="a6687-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="a6687-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="a6687-310">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="a6687-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-311">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="a6687-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="a6687-p110">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="a6687-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="a6687-p111">Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="a6687-p111">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="a6687-p112">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="a6687-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="a6687-319">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="a6687-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-320">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-320">Parameters</span></span>

|<span data-ttu-id="a6687-321">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-321">Name</span></span>| <span data-ttu-id="a6687-322">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-322">Type</span></span>| <span data-ttu-id="a6687-323">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="a6687-324">Object</span><span class="sxs-lookup"><span data-stu-id="a6687-324">Object</span></span> | <span data-ttu-id="a6687-325">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a6687-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="a6687-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="a6687-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="a6687-p113">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="a6687-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="a6687-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="a6687-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="a6687-p114">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="a6687-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="a6687-332">Date</span><span class="sxs-lookup"><span data-stu-id="a6687-332">Date</span></span> | <span data-ttu-id="a6687-333">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a6687-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="a6687-334">Date</span><span class="sxs-lookup"><span data-stu-id="a6687-334">Date</span></span> | <span data-ttu-id="a6687-335">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="a6687-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="a6687-336">String</span><span class="sxs-lookup"><span data-stu-id="a6687-336">String</span></span> | <span data-ttu-id="a6687-p115">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="a6687-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="a6687-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="a6687-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="a6687-p116">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="a6687-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="a6687-342">String</span><span class="sxs-lookup"><span data-stu-id="a6687-342">String</span></span> | <span data-ttu-id="a6687-p117">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="a6687-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="a6687-345">String</span><span class="sxs-lookup"><span data-stu-id="a6687-345">String</span></span> | <span data-ttu-id="a6687-p118">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="a6687-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a6687-348">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-348">Requirements</span></span>

|<span data-ttu-id="a6687-349">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-349">Requirement</span></span>| <span data-ttu-id="a6687-350">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-351">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-352">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-352">1.0</span></span>|
|[<span data-ttu-id="a6687-353">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-354">ReadItem</span></span>|
|[<span data-ttu-id="a6687-355">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-356">Lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6687-357">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="a6687-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a6687-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="a6687-359">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="a6687-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="a6687-p119">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="a6687-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="a6687-362">Vous pouvez transmettre le jeton et soit un identificateur de pièce jointe, soit un identificateur d’élément à un système tiers.</span><span class="sxs-lookup"><span data-stu-id="a6687-362">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="a6687-363">Le système tiers utilise le jeton comme jeton d’autorisation du support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) de services Web Exchange (EWS) ou de [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour renvoyer une pièce jointe ou un élément.</span><span class="sxs-lookup"><span data-stu-id="a6687-363">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="a6687-364">Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="a6687-364">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="a6687-365">L’appel `getCallbackTokenAsync` de la méthode en mode lecture requiert un niveau d’autorisation minimal de **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="a6687-365">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="a6687-366">Pour `getCallbackTokenAsync` appeler en mode composition, vous devez avoir enregistré l’élément.</span><span class="sxs-lookup"><span data-stu-id="a6687-366">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="a6687-367">La [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) méthode requiert un niveau d’autorisation minimum de **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="a6687-367">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-368">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-368">Parameters</span></span>

|<span data-ttu-id="a6687-369">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-369">Name</span></span>| <span data-ttu-id="a6687-370">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-370">Type</span></span>| <span data-ttu-id="a6687-371">Attributs</span><span class="sxs-lookup"><span data-stu-id="a6687-371">Attributes</span></span>| <span data-ttu-id="a6687-372">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a6687-373">fonction</span><span class="sxs-lookup"><span data-stu-id="a6687-373">function</span></span>||<span data-ttu-id="a6687-374">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6687-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6687-375">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6687-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="a6687-376">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="a6687-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="a6687-377">Objet</span><span class="sxs-lookup"><span data-stu-id="a6687-377">Object</span></span>| <span data-ttu-id="a6687-378">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6687-378">&lt;optional&gt;</span></span>|<span data-ttu-id="a6687-379">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a6687-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6687-380">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a6687-380">Errors</span></span>

|<span data-ttu-id="a6687-381">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a6687-381">Error code</span></span>|<span data-ttu-id="a6687-382">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="a6687-383">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="a6687-383">The request has failed.</span></span> <span data-ttu-id="a6687-384">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="a6687-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="a6687-385">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="a6687-385">The Exchange server returned an error.</span></span> <span data-ttu-id="a6687-386">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="a6687-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="a6687-387">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="a6687-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="a6687-388">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="a6687-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-389">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-389">Requirements</span></span>

|<span data-ttu-id="a6687-390">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-390">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="a6687-391">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-391">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-392">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-392">1.0</span></span> | <span data-ttu-id="a6687-393">1.3</span><span class="sxs-lookup"><span data-stu-id="a6687-393">1.3</span></span> |
|[<span data-ttu-id="a6687-394">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-395">ReadItem</span></span> | <span data-ttu-id="a6687-396">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-396">ReadItem</span></span> |
|[<span data-ttu-id="a6687-397">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-397">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-398">Lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-398">Read</span></span> | <span data-ttu-id="a6687-399">Composition</span><span class="sxs-lookup"><span data-stu-id="a6687-399">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="a6687-400">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-400">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="a6687-401">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a6687-401">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="a6687-402">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="a6687-402">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="a6687-403">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="a6687-403">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-404">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-404">Parameters</span></span>

|<span data-ttu-id="a6687-405">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-405">Name</span></span>| <span data-ttu-id="a6687-406">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-406">Type</span></span>| <span data-ttu-id="a6687-407">Attributs</span><span class="sxs-lookup"><span data-stu-id="a6687-407">Attributes</span></span>| <span data-ttu-id="a6687-408">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-408">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="a6687-409">fonction</span><span class="sxs-lookup"><span data-stu-id="a6687-409">function</span></span>||<span data-ttu-id="a6687-410">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6687-410">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6687-411">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6687-411">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="a6687-412">En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="a6687-412">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="a6687-413">Objet</span><span class="sxs-lookup"><span data-stu-id="a6687-413">Object</span></span>| <span data-ttu-id="a6687-414">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6687-414">&lt;optional&gt;</span></span>|<span data-ttu-id="a6687-415">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a6687-415">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="a6687-416">Erreurs</span><span class="sxs-lookup"><span data-stu-id="a6687-416">Errors</span></span>

|<span data-ttu-id="a6687-417">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="a6687-417">Error code</span></span>|<span data-ttu-id="a6687-418">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-418">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="a6687-419">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="a6687-419">The request has failed.</span></span> <span data-ttu-id="a6687-420">Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="a6687-420">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="a6687-421">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="a6687-421">The Exchange server returned an error.</span></span> <span data-ttu-id="a6687-422">Pour plus d’informations, veuillez consulter l’objet de diagnostics.</span><span class="sxs-lookup"><span data-stu-id="a6687-422">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="a6687-423">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="a6687-423">The user is no longer connected to the network.</span></span> <span data-ttu-id="a6687-424">Veuillez vérifier la connexion réseau et réessayer.</span><span class="sxs-lookup"><span data-stu-id="a6687-424">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-425">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-425">Requirements</span></span>

|<span data-ttu-id="a6687-426">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-426">Requirement</span></span>| <span data-ttu-id="a6687-427">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-427">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-428">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-428">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-429">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-429">1.0</span></span>|
|[<span data-ttu-id="a6687-430">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-430">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-431">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a6687-431">ReadItem</span></span>|
|[<span data-ttu-id="a6687-432">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-432">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-433">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-433">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6687-434">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-434">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="a6687-435">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="a6687-435">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="a6687-436">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a6687-436">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-437">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="a6687-437">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="a6687-438">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="a6687-438">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="a6687-439">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="a6687-439">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="a6687-440">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a6687-440">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="a6687-441">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="a6687-441">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="a6687-442">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="a6687-442">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="a6687-443">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="a6687-443">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="a6687-444">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="a6687-444">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="a6687-p129">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="a6687-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="a6687-447">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="a6687-447">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="a6687-448">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="a6687-448">Version differences</span></span>

<span data-ttu-id="a6687-449">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="a6687-449">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="a6687-p130">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="a6687-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="a6687-453">Paramètres</span><span class="sxs-lookup"><span data-stu-id="a6687-453">Parameters</span></span>

|<span data-ttu-id="a6687-454">Nom</span><span class="sxs-lookup"><span data-stu-id="a6687-454">Name</span></span>| <span data-ttu-id="a6687-455">Type</span><span class="sxs-lookup"><span data-stu-id="a6687-455">Type</span></span>| <span data-ttu-id="a6687-456">Attributs</span><span class="sxs-lookup"><span data-stu-id="a6687-456">Attributes</span></span>| <span data-ttu-id="a6687-457">Description</span><span class="sxs-lookup"><span data-stu-id="a6687-457">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="a6687-458">String</span><span class="sxs-lookup"><span data-stu-id="a6687-458">String</span></span>||<span data-ttu-id="a6687-459">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="a6687-459">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="a6687-460">function</span><span class="sxs-lookup"><span data-stu-id="a6687-460">function</span></span>||<span data-ttu-id="a6687-461">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a6687-461">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="a6687-462">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="a6687-462">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="a6687-463">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="a6687-463">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="a6687-464">Objet</span><span class="sxs-lookup"><span data-stu-id="a6687-464">Object</span></span>| <span data-ttu-id="a6687-465">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="a6687-465">&lt;optional&gt;</span></span>|<span data-ttu-id="a6687-466">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a6687-466">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a6687-467">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a6687-467">Requirements</span></span>

|<span data-ttu-id="a6687-468">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a6687-468">Requirement</span></span>| <span data-ttu-id="a6687-469">Valeur</span><span class="sxs-lookup"><span data-stu-id="a6687-469">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6687-470">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a6687-470">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a6687-471">1.0</span><span class="sxs-lookup"><span data-stu-id="a6687-471">1.0</span></span>|
|[<span data-ttu-id="a6687-472">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="a6687-472">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a6687-473">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="a6687-473">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="a6687-474">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a6687-474">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a6687-475">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a6687-475">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6687-476">Exemple</span><span class="sxs-lookup"><span data-stu-id="a6687-476">Example</span></span>

<span data-ttu-id="a6687-477">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="a6687-477">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
