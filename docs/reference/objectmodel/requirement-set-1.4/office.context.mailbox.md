---
title: Office. Context. Mailbox-ensemble de conditions requises 1,4
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 66ae7cb05ac56224fd7461c5c29587e21a24020a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696210"
---
# <a name="mailbox"></a><span data-ttu-id="e3e8d-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="e3e8d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="e3e8d-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="e3e8d-104">Permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="e3e8d-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-105">Requirements</span></span>

|<span data-ttu-id="e3e8d-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-106">Requirement</span></span>| <span data-ttu-id="e3e8d-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-109">1.0</span></span>|
|[<span data-ttu-id="e3e8d-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="e3e8d-111">Restricted</span></span>|
|[<span data-ttu-id="e3e8d-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e3e8d-114">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="e3e8d-114">Members and methods</span></span>

| <span data-ttu-id="e3e8d-115">Membre</span><span class="sxs-lookup"><span data-stu-id="e3e8d-115">Member</span></span> | <span data-ttu-id="e3e8d-116">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e3e8d-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="e3e8d-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="e3e8d-118">Membre</span><span class="sxs-lookup"><span data-stu-id="e3e8d-118">Member</span></span> |
| [<span data-ttu-id="e3e8d-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="e3e8d-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="e3e8d-120">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-120">Method</span></span> |
| [<span data-ttu-id="e3e8d-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e3e8d-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="e3e8d-122">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-122">Method</span></span> |
| [<span data-ttu-id="e3e8d-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="e3e8d-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="e3e8d-124">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-124">Method</span></span> |
| [<span data-ttu-id="e3e8d-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="e3e8d-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="e3e8d-126">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-126">Method</span></span> |
| [<span data-ttu-id="e3e8d-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e3e8d-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="e3e8d-128">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-128">Method</span></span> |
| [<span data-ttu-id="e3e8d-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="e3e8d-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="e3e8d-130">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-130">Method</span></span> |
| [<span data-ttu-id="e3e8d-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="e3e8d-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="e3e8d-132">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-132">Method</span></span> |
| [<span data-ttu-id="e3e8d-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e3e8d-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="e3e8d-134">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-134">Method</span></span> |
| [<span data-ttu-id="e3e8d-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="e3e8d-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="e3e8d-136">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-136">Method</span></span> |
| [<span data-ttu-id="e3e8d-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="e3e8d-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="e3e8d-138">Méthode</span><span class="sxs-lookup"><span data-stu-id="e3e8d-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e3e8d-139">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="e3e8d-139">Namespaces</span></span>

<span data-ttu-id="e3e8d-140">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="e3e8d-141">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="e3e8d-142">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="e3e8d-143">Membres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="e3e8d-144">ewsUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="e3e8d-144">ewsUrl: String</span></span>

<span data-ttu-id="e3e8d-145">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-145">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="e3e8d-146">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-146">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-147">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3e8d-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e3e8d-150">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="e3e8d-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="e3e8d-153">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-153">Type</span></span>

*   <span data-ttu-id="e3e8d-154">String</span><span class="sxs-lookup"><span data-stu-id="e3e8d-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e3e8d-155">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-155">Requirements</span></span>

|<span data-ttu-id="e3e8d-156">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-156">Requirement</span></span>| <span data-ttu-id="e3e8d-157">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-158">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-159">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-159">1.0</span></span>|
|[<span data-ttu-id="e3e8d-160">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-161">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="e3e8d-164">Méthodes</span><span class="sxs-lookup"><span data-stu-id="e3e8d-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="e3e8d-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e3e8d-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e3e8d-166">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-167">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3e8d-p104">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-170">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-170">Parameters</span></span>

|<span data-ttu-id="e3e8d-171">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-171">Name</span></span>| <span data-ttu-id="e3e8d-172">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-172">Type</span></span>| <span data-ttu-id="e3e8d-173">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3e8d-174">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3e8d-174">String</span></span>|<span data-ttu-id="e3e8d-175">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="e3e8d-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="e3e8d-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e3e8d-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="e3e8d-177">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-178">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-178">Requirements</span></span>

|<span data-ttu-id="e3e8d-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-179">Requirement</span></span>| <span data-ttu-id="e3e8d-180">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-181">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-182">1.3</span><span class="sxs-lookup"><span data-stu-id="e3e8d-182">1.3</span></span>|
|[<span data-ttu-id="e3e8d-183">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-184">Restreinte</span><span class="sxs-lookup"><span data-stu-id="e3e8d-184">Restricted</span></span>|
|[<span data-ttu-id="e3e8d-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-186">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3e8d-187">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="e3e8d-187">Returns:</span></span>

<span data-ttu-id="e3e8d-188">Type : String</span><span class="sxs-lookup"><span data-stu-id="e3e8d-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e3e8d-189">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="e3e8d-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="e3e8d-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="e3e8d-191">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="e3e8d-192">Une application de messagerie pour Outlook sur un ordinateur de bureau ou sur le Web peut utiliser différents fuseaux horaires pour les dates et les heures.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-192">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="e3e8d-193">Outlook sur un ordinateur de bureau utilise le fuseau horaire de l’ordinateur client; Outlook sur le Web utilise le fuseau horaire défini dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-193">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="e3e8d-194">Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-194">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="e3e8d-195">Si l’application de messagerie est en cours d’exécution dans Outlook sur un `convertToLocalClientTime` client de bureau, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-195">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="e3e8d-196">Si l’application de messagerie est en cours d’exécution dans Outlook sur `convertToLocalClientTime` le Web, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire spécifié dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-196">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-197">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-197">Parameters</span></span>

|<span data-ttu-id="e3e8d-198">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-198">Name</span></span>| <span data-ttu-id="e3e8d-199">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-199">Type</span></span>| <span data-ttu-id="e3e8d-200">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="e3e8d-201">Date</span><span class="sxs-lookup"><span data-stu-id="e3e8d-201">Date</span></span>|<span data-ttu-id="e3e8d-202">Objet Date</span><span class="sxs-lookup"><span data-stu-id="e3e8d-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-203">Requirements</span></span>

|<span data-ttu-id="e3e8d-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-204">Requirement</span></span>| <span data-ttu-id="e3e8d-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-207">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-207">1.0</span></span>|
|[<span data-ttu-id="e3e8d-208">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-209">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-211">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3e8d-212">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="e3e8d-212">Returns:</span></span>

<span data-ttu-id="e3e8d-213">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="e3e8d-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="e3e8d-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="e3e8d-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="e3e8d-215">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-216">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3e8d-p107">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-219">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-219">Parameters</span></span>

|<span data-ttu-id="e3e8d-220">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-220">Name</span></span>| <span data-ttu-id="e3e8d-221">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-221">Type</span></span>| <span data-ttu-id="e3e8d-222">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3e8d-223">String</span><span class="sxs-lookup"><span data-stu-id="e3e8d-223">String</span></span>|<span data-ttu-id="e3e8d-224">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="e3e8d-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="e3e8d-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="e3e8d-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="e3e8d-226">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-227">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-227">Requirements</span></span>

|<span data-ttu-id="e3e8d-228">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-228">Requirement</span></span>| <span data-ttu-id="e3e8d-229">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-230">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-231">1.3</span><span class="sxs-lookup"><span data-stu-id="e3e8d-231">1.3</span></span>|
|[<span data-ttu-id="e3e8d-232">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-233">Restreinte</span><span class="sxs-lookup"><span data-stu-id="e3e8d-233">Restricted</span></span>|
|[<span data-ttu-id="e3e8d-234">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-235">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3e8d-236">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="e3e8d-236">Returns:</span></span>

<span data-ttu-id="e3e8d-237">Type : String</span><span class="sxs-lookup"><span data-stu-id="e3e8d-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="e3e8d-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="e3e8d-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="e3e8d-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="e3e8d-240">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="e3e8d-241">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-242">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-242">Parameters</span></span>

|<span data-ttu-id="e3e8d-243">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-243">Name</span></span>| <span data-ttu-id="e3e8d-244">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-244">Type</span></span>| <span data-ttu-id="e3e8d-245">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="e3e8d-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="e3e8d-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)|<span data-ttu-id="e3e8d-247">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-248">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-248">Requirements</span></span>

|<span data-ttu-id="e3e8d-249">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-249">Requirement</span></span>| <span data-ttu-id="e3e8d-250">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-251">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-252">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-252">1.0</span></span>|
|[<span data-ttu-id="e3e8d-253">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-254">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-255">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-256">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="e3e8d-257">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="e3e8d-257">Returns:</span></span>

<span data-ttu-id="e3e8d-258">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="e3e8d-259">Type: date</span><span class="sxs-lookup"><span data-stu-id="e3e8d-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="e3e8d-260">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="e3e8d-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e3e8d-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="e3e8d-262">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-263">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3e8d-264">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e3e8d-265">Dans Outlook sur Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-265">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="e3e8d-266">En effet, dans Outlook sur Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-266">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="e3e8d-267">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32KO nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="e3e8d-268">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-269">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-269">Parameters</span></span>

|<span data-ttu-id="e3e8d-270">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-270">Name</span></span>| <span data-ttu-id="e3e8d-271">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-271">Type</span></span>| <span data-ttu-id="e3e8d-272">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3e8d-273">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3e8d-273">String</span></span>|<span data-ttu-id="e3e8d-274">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-275">Requirements</span></span>

|<span data-ttu-id="e3e8d-276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-276">Requirement</span></span>| <span data-ttu-id="e3e8d-277">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-279">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-279">1.0</span></span>|
|[<span data-ttu-id="e3e8d-280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-281">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-283">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3e8d-284">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="e3e8d-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="e3e8d-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="e3e8d-286">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-287">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3e8d-288">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="e3e8d-289">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32 Ko nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="e3e8d-290">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="e3e8d-p109">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-293">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-293">Parameters</span></span>

|<span data-ttu-id="e3e8d-294">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-294">Name</span></span>| <span data-ttu-id="e3e8d-295">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-295">Type</span></span>| <span data-ttu-id="e3e8d-296">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="e3e8d-297">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3e8d-297">String</span></span>|<span data-ttu-id="e3e8d-298">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-299">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-299">Requirements</span></span>

|<span data-ttu-id="e3e8d-300">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-300">Requirement</span></span>| <span data-ttu-id="e3e8d-301">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-302">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-303">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-303">1.0</span></span>|
|[<span data-ttu-id="e3e8d-304">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-305">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-306">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-307">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3e8d-308">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="e3e8d-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="e3e8d-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="e3e8d-310">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-311">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="e3e8d-p110">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="e3e8d-314">Dans Outlook sur le Web et les appareils mobiles, cette méthode affiche toujours un formulaire avec un champ participants.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-314">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="e3e8d-315">Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-315">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="e3e8d-316">Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-316">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="e3e8d-p112">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="e3e8d-319">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-320">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-320">Parameters</span></span>

|<span data-ttu-id="e3e8d-321">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-321">Name</span></span>| <span data-ttu-id="e3e8d-322">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-322">Type</span></span>| <span data-ttu-id="e3e8d-323">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="e3e8d-324">Object</span><span class="sxs-lookup"><span data-stu-id="e3e8d-324">Object</span></span> | <span data-ttu-id="e3e8d-325">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="e3e8d-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="e3e8d-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="e3e8d-p113">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="e3e8d-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="e3e8d-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="e3e8d-p114">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="e3e8d-332">Date</span><span class="sxs-lookup"><span data-stu-id="e3e8d-332">Date</span></span> | <span data-ttu-id="e3e8d-333">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="e3e8d-334">Date</span><span class="sxs-lookup"><span data-stu-id="e3e8d-334">Date</span></span> | <span data-ttu-id="e3e8d-335">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="e3e8d-336">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3e8d-336">String</span></span> | <span data-ttu-id="e3e8d-p115">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="e3e8d-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="e3e8d-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="e3e8d-p116">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="e3e8d-342">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3e8d-342">String</span></span> | <span data-ttu-id="e3e8d-p117">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="e3e8d-345">String</span><span class="sxs-lookup"><span data-stu-id="e3e8d-345">String</span></span> | <span data-ttu-id="e3e8d-p118">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e3e8d-348">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-348">Requirements</span></span>

|<span data-ttu-id="e3e8d-349">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-349">Requirement</span></span>| <span data-ttu-id="e3e8d-350">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-351">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-352">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-352">1.0</span></span>|
|[<span data-ttu-id="e3e8d-353">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-354">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-355">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-356">Lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3e8d-357">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="e3e8d-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e3e8d-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e3e8d-359">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="e3e8d-p119">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="e3e8d-p120">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="e3e8d-365">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-365">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="e3e8d-p121">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-368">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-368">Parameters</span></span>

|<span data-ttu-id="e3e8d-369">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-369">Name</span></span>| <span data-ttu-id="e3e8d-370">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-370">Type</span></span>| <span data-ttu-id="e3e8d-371">Attributs</span><span class="sxs-lookup"><span data-stu-id="e3e8d-371">Attributes</span></span>| <span data-ttu-id="e3e8d-372">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e3e8d-373">fonction</span><span class="sxs-lookup"><span data-stu-id="e3e8d-373">function</span></span>||<span data-ttu-id="e3e8d-374">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e3e8d-375">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e3e8d-376">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e3e8d-377">Objet</span><span class="sxs-lookup"><span data-stu-id="e3e8d-377">Object</span></span>| <span data-ttu-id="e3e8d-378">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e3e8d-378">&lt;optional&gt;</span></span>|<span data-ttu-id="e3e8d-379">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e3e8d-380">Erreurs</span><span class="sxs-lookup"><span data-stu-id="e3e8d-380">Errors</span></span>

|<span data-ttu-id="e3e8d-381">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-381">Error code</span></span>|<span data-ttu-id="e3e8d-382">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e3e8d-383">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-383">The request has failed.</span></span> <span data-ttu-id="e3e8d-384">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e3e8d-385">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-385">The Exchange server returned an error.</span></span> <span data-ttu-id="e3e8d-386">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e3e8d-387">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="e3e8d-388">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-389">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-389">Requirements</span></span>

|<span data-ttu-id="e3e8d-390">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-390">Requirement</span></span>| <span data-ttu-id="e3e8d-391">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-392">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-393">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-393">1.0</span></span>|
|[<span data-ttu-id="e3e8d-394">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-395">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-396">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-397">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-397">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3e8d-398">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-398">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="e3e8d-399">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e3e8d-399">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="e3e8d-400">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-400">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="e3e8d-401">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-401">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-402">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-402">Parameters</span></span>

|<span data-ttu-id="e3e8d-403">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-403">Name</span></span>| <span data-ttu-id="e3e8d-404">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-404">Type</span></span>| <span data-ttu-id="e3e8d-405">Attributs</span><span class="sxs-lookup"><span data-stu-id="e3e8d-405">Attributes</span></span>| <span data-ttu-id="e3e8d-406">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-406">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="e3e8d-407">fonction</span><span class="sxs-lookup"><span data-stu-id="e3e8d-407">function</span></span>||<span data-ttu-id="e3e8d-408">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-408">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e3e8d-409">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-409">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="e3e8d-410">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-410">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="e3e8d-411">Objet</span><span class="sxs-lookup"><span data-stu-id="e3e8d-411">Object</span></span>| <span data-ttu-id="e3e8d-412">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e3e8d-412">&lt;optional&gt;</span></span>|<span data-ttu-id="e3e8d-413">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-413">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="e3e8d-414">Erreurs</span><span class="sxs-lookup"><span data-stu-id="e3e8d-414">Errors</span></span>

|<span data-ttu-id="e3e8d-415">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-415">Error code</span></span>|<span data-ttu-id="e3e8d-416">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-416">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="e3e8d-417">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-417">The request has failed.</span></span> <span data-ttu-id="e3e8d-418">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-418">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="e3e8d-419">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-419">The Exchange server returned an error.</span></span> <span data-ttu-id="e3e8d-420">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-420">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="e3e8d-421">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-421">The user is no longer connected to the network.</span></span> <span data-ttu-id="e3e8d-422">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-422">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-423">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-423">Requirements</span></span>

|<span data-ttu-id="e3e8d-424">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-424">Requirement</span></span>| <span data-ttu-id="e3e8d-425">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-426">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-427">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-427">1.0</span></span>|
|[<span data-ttu-id="e3e8d-428">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="e3e8d-429">ReadItem</span></span>|
|[<span data-ttu-id="e3e8d-430">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-431">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3e8d-432">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-432">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="e3e8d-433">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="e3e8d-433">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="e3e8d-434">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-434">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-435">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="e3e8d-435">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="e3e8d-436">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="e3e8d-436">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="e3e8d-437">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-437">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="e3e8d-438">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-438">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="e3e8d-439">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-439">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="e3e8d-440">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-440">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="e3e8d-441">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-441">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="e3e8d-442">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-442">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="e3e8d-p129">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="e3e8d-445">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-445">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="e3e8d-446">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="e3e8d-446">Version differences</span></span>

<span data-ttu-id="e3e8d-447">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-447">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="e3e8d-p130">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="e3e8d-451">Paramètres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-451">Parameters</span></span>

|<span data-ttu-id="e3e8d-452">Nom</span><span class="sxs-lookup"><span data-stu-id="e3e8d-452">Name</span></span>| <span data-ttu-id="e3e8d-453">Type</span><span class="sxs-lookup"><span data-stu-id="e3e8d-453">Type</span></span>| <span data-ttu-id="e3e8d-454">Attributs</span><span class="sxs-lookup"><span data-stu-id="e3e8d-454">Attributes</span></span>| <span data-ttu-id="e3e8d-455">Description</span><span class="sxs-lookup"><span data-stu-id="e3e8d-455">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="e3e8d-456">String</span><span class="sxs-lookup"><span data-stu-id="e3e8d-456">String</span></span>||<span data-ttu-id="e3e8d-457">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-457">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="e3e8d-458">function</span><span class="sxs-lookup"><span data-stu-id="e3e8d-458">function</span></span>||<span data-ttu-id="e3e8d-459">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="e3e8d-459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="e3e8d-460">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-460">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="e3e8d-461">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-461">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="e3e8d-462">Objet</span><span class="sxs-lookup"><span data-stu-id="e3e8d-462">Object</span></span>| <span data-ttu-id="e3e8d-463">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="e3e8d-463">&lt;optional&gt;</span></span>|<span data-ttu-id="e3e8d-464">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-464">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3e8d-465">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3e8d-465">Requirements</span></span>

|<span data-ttu-id="e3e8d-466">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3e8d-466">Requirement</span></span>| <span data-ttu-id="e3e8d-467">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3e8d-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3e8d-468">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3e8d-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3e8d-469">1.0</span><span class="sxs-lookup"><span data-stu-id="e3e8d-469">1.0</span></span>|
|[<span data-ttu-id="e3e8d-470">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e3e8d-470">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="e3e8d-471">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="e3e8d-471">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="e3e8d-472">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3e8d-472">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3e8d-473">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3e8d-473">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e3e8d-474">Exemple</span><span class="sxs-lookup"><span data-stu-id="e3e8d-474">Example</span></span>

<span data-ttu-id="e3e8d-475">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="e3e8d-475">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
