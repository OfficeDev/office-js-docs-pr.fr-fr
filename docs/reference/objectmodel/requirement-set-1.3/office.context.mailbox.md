---
title: Office. Context. Mailbox-ensemble de conditions requises 1,3
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 05b7d82e036cc29526c18bf97c6a1472778c1959
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696231"
---
# <a name="mailbox"></a><span data-ttu-id="11a8b-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="11a8b-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="11a8b-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="11a8b-104">Permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="11a8b-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="11a8b-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-105">Requirements</span></span>

|<span data-ttu-id="11a8b-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-106">Requirement</span></span>| <span data-ttu-id="11a8b-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-109">1.0</span></span>|
|[<span data-ttu-id="11a8b-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="11a8b-111">Restricted</span></span>|
|[<span data-ttu-id="11a8b-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-113">Compose or Read</span></span>|

<span data-ttu-id="11a8b-114">| [ewsUrl](#ewsurl-string) | Membre | | [convertToEwsId](#converttoewsiditemid-restversion--string) | Méthode | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Méthode | | [convertToRestId](#converttorestiditemid-restversion--string) | Méthode | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Méthode | | [displayAppointmentForm](#displayappointmentformitemid) | Méthode | | [displayMessageForm](#displaymessageformitemid) | Méthode | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Méthode | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Méthode | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Méthode | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Méthode |</span><span class="sxs-lookup"><span data-stu-id="11a8b-114">| [ewsUrl](#ewsurl-string) | Member | | [convertToEwsId](#converttoewsiditemid-restversion--string) | Method | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Method | | [convertToRestId](#converttorestiditemid-restversion--string) | Method | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Method | | [displayAppointmentForm](#displayappointmentformitemid) | Method | | [displayMessageForm](#displaymessageformitemid) | Method | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Method | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Method | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Method | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Method |</span></span>

### <a name="namespaces"></a><span data-ttu-id="11a8b-115">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="11a8b-115">Namespaces</span></span>

<span data-ttu-id="11a8b-116">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="11a8b-116">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="11a8b-117">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="11a8b-117">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="11a8b-118">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="11a8b-118">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="11a8b-119">Membres</span><span class="sxs-lookup"><span data-stu-id="11a8b-119">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="11a8b-120">ewsUrl: chaîne</span><span class="sxs-lookup"><span data-stu-id="11a8b-120">ewsUrl: String</span></span>

<span data-ttu-id="11a8b-121">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie.</span><span class="sxs-lookup"><span data-stu-id="11a8b-121">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="11a8b-122">Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="11a8b-122">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-123">Ce membre n’est pas pris en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="11a8b-123">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="11a8b-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="11a8b-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="11a8b-126">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="11a8b-126">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="11a8b-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="11a8b-129">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-129">Type</span></span>

*   <span data-ttu-id="11a8b-130">String</span><span class="sxs-lookup"><span data-stu-id="11a8b-130">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11a8b-131">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-131">Requirements</span></span>

|<span data-ttu-id="11a8b-132">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-132">Requirement</span></span>| <span data-ttu-id="11a8b-133">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-133">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-134">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-135">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-135">1.0</span></span>|
|[<span data-ttu-id="11a8b-136">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-136">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-137">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-138">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-138">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-139">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-139">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="11a8b-140">Méthodes</span><span class="sxs-lookup"><span data-stu-id="11a8b-140">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="11a8b-141">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="11a8b-141">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="11a8b-142">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="11a8b-142">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-143">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="11a8b-143">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="11a8b-p104">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-146">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-146">Parameters</span></span>

|<span data-ttu-id="11a8b-147">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-147">Name</span></span>| <span data-ttu-id="11a8b-148">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-148">Type</span></span>| <span data-ttu-id="11a8b-149">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-149">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="11a8b-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="11a8b-150">String</span></span>|<span data-ttu-id="11a8b-151">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="11a8b-151">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="11a8b-152">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="11a8b-152">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="11a8b-153">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="11a8b-153">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-154">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-154">Requirements</span></span>

|<span data-ttu-id="11a8b-155">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-155">Requirement</span></span>| <span data-ttu-id="11a8b-156">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-157">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-158">1.3</span><span class="sxs-lookup"><span data-stu-id="11a8b-158">1.3</span></span>|
|[<span data-ttu-id="11a8b-159">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-159">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-160">Restreinte</span><span class="sxs-lookup"><span data-stu-id="11a8b-160">Restricted</span></span>|
|[<span data-ttu-id="11a8b-161">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-162">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-162">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="11a8b-163">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="11a8b-163">Returns:</span></span>

<span data-ttu-id="11a8b-164">Type : String</span><span class="sxs-lookup"><span data-stu-id="11a8b-164">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="11a8b-165">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-165">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="11a8b-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="11a8b-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="11a8b-167">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="11a8b-167">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="11a8b-168">Une application de messagerie pour Outlook sur un ordinateur de bureau ou sur le Web peut utiliser différents fuseaux horaires pour les dates et les heures.</span><span class="sxs-lookup"><span data-stu-id="11a8b-168">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="11a8b-169">Outlook sur un ordinateur de bureau utilise le fuseau horaire de l’ordinateur client; Outlook sur le Web utilise le fuseau horaire défini dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="11a8b-169">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="11a8b-170">Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="11a8b-170">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="11a8b-171">Si l’application de messagerie est en cours d’exécution dans Outlook sur un `convertToLocalClientTime` client de bureau, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire de l’ordinateur client.</span><span class="sxs-lookup"><span data-stu-id="11a8b-171">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="11a8b-172">Si l’application de messagerie est en cours d’exécution dans Outlook sur `convertToLocalClientTime` le Web, la méthode renvoie un objet Dictionary dont les valeurs sont définies sur le fuseau horaire spécifié dans le centre d’administration Exchange.</span><span class="sxs-lookup"><span data-stu-id="11a8b-172">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-173">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-173">Parameters</span></span>

|<span data-ttu-id="11a8b-174">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-174">Name</span></span>| <span data-ttu-id="11a8b-175">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-175">Type</span></span>| <span data-ttu-id="11a8b-176">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-176">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="11a8b-177">Date</span><span class="sxs-lookup"><span data-stu-id="11a8b-177">Date</span></span>|<span data-ttu-id="11a8b-178">Objet Date</span><span class="sxs-lookup"><span data-stu-id="11a8b-178">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-179">Requirements</span></span>

|<span data-ttu-id="11a8b-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-180">Requirement</span></span>| <span data-ttu-id="11a8b-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-183">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-183">1.0</span></span>|
|[<span data-ttu-id="11a8b-184">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-185">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-187">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="11a8b-188">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="11a8b-188">Returns:</span></span>

<span data-ttu-id="11a8b-189">Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="11a8b-189">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="11a8b-190">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="11a8b-190">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="11a8b-191">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="11a8b-191">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-192">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="11a8b-192">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="11a8b-p107">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-195">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-195">Parameters</span></span>

|<span data-ttu-id="11a8b-196">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-196">Name</span></span>| <span data-ttu-id="11a8b-197">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-197">Type</span></span>| <span data-ttu-id="11a8b-198">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-198">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="11a8b-199">String</span><span class="sxs-lookup"><span data-stu-id="11a8b-199">String</span></span>|<span data-ttu-id="11a8b-200">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="11a8b-200">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="11a8b-201">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="11a8b-201">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="11a8b-202">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="11a8b-202">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-203">Requirements</span></span>

|<span data-ttu-id="11a8b-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-204">Requirement</span></span>| <span data-ttu-id="11a8b-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-207">1.3</span><span class="sxs-lookup"><span data-stu-id="11a8b-207">1.3</span></span>|
|[<span data-ttu-id="11a8b-208">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-209">Restreinte</span><span class="sxs-lookup"><span data-stu-id="11a8b-209">Restricted</span></span>|
|[<span data-ttu-id="11a8b-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-211">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="11a8b-212">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="11a8b-212">Returns:</span></span>

<span data-ttu-id="11a8b-213">Type : String</span><span class="sxs-lookup"><span data-stu-id="11a8b-213">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="11a8b-214">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-214">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="11a8b-215">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="11a8b-215">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="11a8b-216">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="11a8b-216">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="11a8b-217">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="11a8b-217">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-218">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-218">Parameters</span></span>

|<span data-ttu-id="11a8b-219">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-219">Name</span></span>| <span data-ttu-id="11a8b-220">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-220">Type</span></span>| <span data-ttu-id="11a8b-221">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-221">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="11a8b-222">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="11a8b-222">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="11a8b-223">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="11a8b-223">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-224">Requirements</span></span>

|<span data-ttu-id="11a8b-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-225">Requirement</span></span>| <span data-ttu-id="11a8b-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-228">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-228">1.0</span></span>|
|[<span data-ttu-id="11a8b-229">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-230">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-231">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-232">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-232">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="11a8b-233">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="11a8b-233">Returns:</span></span>

<span data-ttu-id="11a8b-234">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="11a8b-234">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="11a8b-235">Type: date</span><span class="sxs-lookup"><span data-stu-id="11a8b-235">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="11a8b-236">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-236">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="11a8b-237">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="11a8b-237">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="11a8b-238">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="11a8b-238">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-239">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="11a8b-239">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="11a8b-240">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="11a8b-240">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="11a8b-241">Dans Outlook sur Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série.</span><span class="sxs-lookup"><span data-stu-id="11a8b-241">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="11a8b-242">En effet, dans Outlook sur Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="11a8b-242">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="11a8b-243">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32KO nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="11a8b-243">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="11a8b-244">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="11a8b-244">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-245">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-245">Parameters</span></span>

|<span data-ttu-id="11a8b-246">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-246">Name</span></span>| <span data-ttu-id="11a8b-247">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-247">Type</span></span>| <span data-ttu-id="11a8b-248">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="11a8b-249">Chaîne</span><span class="sxs-lookup"><span data-stu-id="11a8b-249">String</span></span>|<span data-ttu-id="11a8b-250">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="11a8b-250">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-251">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-251">Requirements</span></span>

|<span data-ttu-id="11a8b-252">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-252">Requirement</span></span>| <span data-ttu-id="11a8b-253">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-254">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-255">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-255">1.0</span></span>|
|[<span data-ttu-id="11a8b-256">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-257">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-258">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-259">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11a8b-260">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-260">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="11a8b-261">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="11a8b-261">displayMessageForm(itemId)</span></span>

<span data-ttu-id="11a8b-262">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="11a8b-262">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-263">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="11a8b-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="11a8b-264">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="11a8b-264">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="11a8b-265">Dans Outlook sur le Web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire est inférieur ou égal à 32 Ko nombre de caractères.</span><span class="sxs-lookup"><span data-stu-id="11a8b-265">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="11a8b-266">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="11a8b-266">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="11a8b-p109">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-269">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-269">Parameters</span></span>

|<span data-ttu-id="11a8b-270">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-270">Name</span></span>| <span data-ttu-id="11a8b-271">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-271">Type</span></span>| <span data-ttu-id="11a8b-272">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="11a8b-273">Chaîne</span><span class="sxs-lookup"><span data-stu-id="11a8b-273">String</span></span>|<span data-ttu-id="11a8b-274">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="11a8b-274">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-275">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-275">Requirements</span></span>

|<span data-ttu-id="11a8b-276">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-276">Requirement</span></span>| <span data-ttu-id="11a8b-277">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-278">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-279">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-279">1.0</span></span>|
|[<span data-ttu-id="11a8b-280">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-281">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-282">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-283">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11a8b-284">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-284">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="11a8b-285">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="11a8b-285">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="11a8b-286">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="11a8b-286">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-287">Cette méthode n’est pas prise en charge dans Outlook sur iOS ou Android.</span><span class="sxs-lookup"><span data-stu-id="11a8b-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="11a8b-p110">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="11a8b-290">Dans Outlook sur le Web et les appareils mobiles, cette méthode affiche toujours un formulaire avec un champ participants.</span><span class="sxs-lookup"><span data-stu-id="11a8b-290">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="11a8b-291">Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**.</span><span class="sxs-lookup"><span data-stu-id="11a8b-291">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="11a8b-292">Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="11a8b-292">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="11a8b-p112">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="11a8b-295">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="11a8b-295">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-296">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-296">Parameters</span></span>

|<span data-ttu-id="11a8b-297">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-297">Name</span></span>| <span data-ttu-id="11a8b-298">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-298">Type</span></span>| <span data-ttu-id="11a8b-299">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-299">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="11a8b-300">Object</span><span class="sxs-lookup"><span data-stu-id="11a8b-300">Object</span></span> | <span data-ttu-id="11a8b-301">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="11a8b-301">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="11a8b-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="11a8b-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="11a8b-p113">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="11a8b-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="11a8b-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="11a8b-p114">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="11a8b-308">Date</span><span class="sxs-lookup"><span data-stu-id="11a8b-308">Date</span></span> | <span data-ttu-id="11a8b-309">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="11a8b-309">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="11a8b-310">Date</span><span class="sxs-lookup"><span data-stu-id="11a8b-310">Date</span></span> | <span data-ttu-id="11a8b-311">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="11a8b-311">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="11a8b-312">Chaîne</span><span class="sxs-lookup"><span data-stu-id="11a8b-312">String</span></span> | <span data-ttu-id="11a8b-p115">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="11a8b-315">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="11a8b-315">Array.&lt;String&gt;</span></span> | <span data-ttu-id="11a8b-p116">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="11a8b-318">Chaîne</span><span class="sxs-lookup"><span data-stu-id="11a8b-318">String</span></span> | <span data-ttu-id="11a8b-p117">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="11a8b-321">String</span><span class="sxs-lookup"><span data-stu-id="11a8b-321">String</span></span> | <span data-ttu-id="11a8b-p118">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="11a8b-324">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-324">Requirements</span></span>

|<span data-ttu-id="11a8b-325">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-325">Requirement</span></span>| <span data-ttu-id="11a8b-326">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-327">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-328">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-328">1.0</span></span>|
|[<span data-ttu-id="11a8b-329">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-330">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-331">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-332">Lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11a8b-333">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-333">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="11a8b-334">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="11a8b-334">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="11a8b-335">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="11a8b-335">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="11a8b-p119">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="11a8b-p120">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="11a8b-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="11a8b-341">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="11a8b-341">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="11a8b-p121">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-344">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-344">Parameters</span></span>

|<span data-ttu-id="11a8b-345">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-345">Name</span></span>| <span data-ttu-id="11a8b-346">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-346">Type</span></span>| <span data-ttu-id="11a8b-347">Attributs</span><span class="sxs-lookup"><span data-stu-id="11a8b-347">Attributes</span></span>| <span data-ttu-id="11a8b-348">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-348">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="11a8b-349">fonction</span><span class="sxs-lookup"><span data-stu-id="11a8b-349">function</span></span>||<span data-ttu-id="11a8b-350">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="11a8b-350">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="11a8b-351">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="11a8b-351">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="11a8b-352">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="11a8b-352">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="11a8b-353">Objet</span><span class="sxs-lookup"><span data-stu-id="11a8b-353">Object</span></span>| <span data-ttu-id="11a8b-354">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="11a8b-354">&lt;optional&gt;</span></span>|<span data-ttu-id="11a8b-355">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="11a8b-355">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="11a8b-356">Erreurs</span><span class="sxs-lookup"><span data-stu-id="11a8b-356">Errors</span></span>

|<span data-ttu-id="11a8b-357">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="11a8b-357">Error code</span></span>|<span data-ttu-id="11a8b-358">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-358">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="11a8b-359">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="11a8b-359">The request has failed.</span></span> <span data-ttu-id="11a8b-360">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="11a8b-360">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="11a8b-361">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="11a8b-361">The Exchange server returned an error.</span></span> <span data-ttu-id="11a8b-362">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="11a8b-362">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="11a8b-363">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="11a8b-363">The user is no longer connected to the network.</span></span> <span data-ttu-id="11a8b-364">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="11a8b-364">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-365">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-365">Requirements</span></span>

|<span data-ttu-id="11a8b-366">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-366">Requirement</span></span>| <span data-ttu-id="11a8b-367">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-368">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-369">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-369">1.0</span></span>|
|[<span data-ttu-id="11a8b-370">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-371">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-372">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-373">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-373">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="11a8b-374">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-374">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="11a8b-375">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="11a8b-375">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="11a8b-376">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="11a8b-376">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="11a8b-377">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="11a8b-377">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-378">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-378">Parameters</span></span>

|<span data-ttu-id="11a8b-379">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-379">Name</span></span>| <span data-ttu-id="11a8b-380">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-380">Type</span></span>| <span data-ttu-id="11a8b-381">Attributs</span><span class="sxs-lookup"><span data-stu-id="11a8b-381">Attributes</span></span>| <span data-ttu-id="11a8b-382">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-382">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="11a8b-383">fonction</span><span class="sxs-lookup"><span data-stu-id="11a8b-383">function</span></span>||<span data-ttu-id="11a8b-384">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="11a8b-384">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="11a8b-385">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="11a8b-385">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="11a8b-386">Si une erreur s’est produite, `asyncResult.error` les `asyncResult.diagnostics` propriétés et peuvent fournir des informations supplémentaires.</span><span class="sxs-lookup"><span data-stu-id="11a8b-386">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="11a8b-387">Objet</span><span class="sxs-lookup"><span data-stu-id="11a8b-387">Object</span></span>| <span data-ttu-id="11a8b-388">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="11a8b-388">&lt;optional&gt;</span></span>|<span data-ttu-id="11a8b-389">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="11a8b-389">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="11a8b-390">Erreurs</span><span class="sxs-lookup"><span data-stu-id="11a8b-390">Errors</span></span>

|<span data-ttu-id="11a8b-391">Code d'erreur</span><span class="sxs-lookup"><span data-stu-id="11a8b-391">Error code</span></span>|<span data-ttu-id="11a8b-392">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-392">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="11a8b-393">La demande a échoué.</span><span class="sxs-lookup"><span data-stu-id="11a8b-393">The request has failed.</span></span> <span data-ttu-id="11a8b-394">Consultez l’objet Diagnostics pour obtenir le code d’erreur HTTP.</span><span class="sxs-lookup"><span data-stu-id="11a8b-394">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="11a8b-395">Le serveur Exchange a renvoyé une erreur.</span><span class="sxs-lookup"><span data-stu-id="11a8b-395">The Exchange server returned an error.</span></span> <span data-ttu-id="11a8b-396">Pour plus d’informations, consultez l’objet Diagnostics.</span><span class="sxs-lookup"><span data-stu-id="11a8b-396">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="11a8b-397">L’utilisateur n’est plus connecté au réseau.</span><span class="sxs-lookup"><span data-stu-id="11a8b-397">The user is no longer connected to the network.</span></span> <span data-ttu-id="11a8b-398">Vérifiez votre connexion réseau, puis réessayez.</span><span class="sxs-lookup"><span data-stu-id="11a8b-398">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-399">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-399">Requirements</span></span>

|<span data-ttu-id="11a8b-400">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-400">Requirement</span></span>| <span data-ttu-id="11a8b-401">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-402">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-403">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-403">1.0</span></span>|
|[<span data-ttu-id="11a8b-404">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11a8b-405">ReadItem</span></span>|
|[<span data-ttu-id="11a8b-406">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-407">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-407">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11a8b-408">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-408">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="11a8b-409">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="11a8b-409">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="11a8b-410">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="11a8b-410">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-411">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="11a8b-411">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="11a8b-412">Dans Outlook sur iOS ou Android</span><span class="sxs-lookup"><span data-stu-id="11a8b-412">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="11a8b-413">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="11a8b-413">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="11a8b-414">Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="11a8b-414">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="11a8b-415">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="11a8b-415">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="11a8b-416">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="11a8b-416">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="11a8b-417">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="11a8b-417">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="11a8b-418">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="11a8b-418">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="11a8b-p129">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="11a8b-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="11a8b-421">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="11a8b-421">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="11a8b-422">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="11a8b-422">Version differences</span></span>

<span data-ttu-id="11a8b-423">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="11a8b-423">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="11a8b-p130">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="11a8b-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="11a8b-427">Paramètres</span><span class="sxs-lookup"><span data-stu-id="11a8b-427">Parameters</span></span>

|<span data-ttu-id="11a8b-428">Nom</span><span class="sxs-lookup"><span data-stu-id="11a8b-428">Name</span></span>| <span data-ttu-id="11a8b-429">Type</span><span class="sxs-lookup"><span data-stu-id="11a8b-429">Type</span></span>| <span data-ttu-id="11a8b-430">Attributs</span><span class="sxs-lookup"><span data-stu-id="11a8b-430">Attributes</span></span>| <span data-ttu-id="11a8b-431">Description</span><span class="sxs-lookup"><span data-stu-id="11a8b-431">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="11a8b-432">String</span><span class="sxs-lookup"><span data-stu-id="11a8b-432">String</span></span>||<span data-ttu-id="11a8b-433">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="11a8b-433">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="11a8b-434">function</span><span class="sxs-lookup"><span data-stu-id="11a8b-434">function</span></span>||<span data-ttu-id="11a8b-435">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="11a8b-435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="11a8b-436">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="11a8b-436">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="11a8b-437">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="11a8b-437">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="11a8b-438">Objet</span><span class="sxs-lookup"><span data-stu-id="11a8b-438">Object</span></span>| <span data-ttu-id="11a8b-439">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="11a8b-439">&lt;optional&gt;</span></span>|<span data-ttu-id="11a8b-440">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="11a8b-440">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="11a8b-441">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="11a8b-441">Requirements</span></span>

|<span data-ttu-id="11a8b-442">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="11a8b-442">Requirement</span></span>| <span data-ttu-id="11a8b-443">Valeur</span><span class="sxs-lookup"><span data-stu-id="11a8b-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="11a8b-444">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="11a8b-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11a8b-445">1.0</span><span class="sxs-lookup"><span data-stu-id="11a8b-445">1.0</span></span>|
|[<span data-ttu-id="11a8b-446">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="11a8b-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11a8b-447">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="11a8b-447">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="11a8b-448">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="11a8b-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11a8b-449">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="11a8b-449">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11a8b-450">Exemple</span><span class="sxs-lookup"><span data-stu-id="11a8b-450">Example</span></span>

<span data-ttu-id="11a8b-451">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="11a8b-451">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
