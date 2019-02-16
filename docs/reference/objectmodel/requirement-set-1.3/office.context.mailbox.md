---
title: Office.context – ensemble de conditions requises 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: edf9bcaa842660007bea85b497308aa7949f8085
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068223"
---
# <a name="mailbox"></a><span data-ttu-id="4159c-102">boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="4159c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="4159c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="4159c-104">Permet d’accéder au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="4159c-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4159c-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-105">Requirements</span></span>

|<span data-ttu-id="4159c-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-106">Requirement</span></span>| <span data-ttu-id="4159c-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-109">1.0</span></span>|
|[<span data-ttu-id="4159c-110">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-111">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4159c-111">Restricted</span></span>|
|[<span data-ttu-id="4159c-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-113">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="4159c-114">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4159c-114">Namespaces</span></span>

<span data-ttu-id="4159c-115">[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4159c-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4159c-116">[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4159c-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4159c-117">[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4159c-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4159c-118">Membres</span><span class="sxs-lookup"><span data-stu-id="4159c-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4159c-119">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="4159c-119">ewsUrl :String</span></span>

<span data-ttu-id="4159c-p101">Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.</span><span class="sxs-lookup"><span data-stu-id="4159c-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-122">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4159c-122">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4159c-p102">La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4159c-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4159c-125">Votre application doit avoir l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler le membre `ewsUrl` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="4159c-125">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="4159c-p103">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4159c-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="4159c-128">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-128">Type</span></span>

*   <span data-ttu-id="4159c-129">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4159c-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4159c-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-130">Requirements</span></span>

|<span data-ttu-id="4159c-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-131">Requirement</span></span>| <span data-ttu-id="4159c-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-134">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-134">1.0</span></span>|
|[<span data-ttu-id="4159c-135">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-135">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-136">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-136">ReadItem</span></span>|
|[<span data-ttu-id="4159c-137">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-137">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-138">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-138">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4159c-139">Méthodes</span><span class="sxs-lookup"><span data-stu-id="4159c-139">Methods</span></span>

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="4159c-140">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4159c-140">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4159c-141">Convertit un ID d’élément mis en forme pour REST au format EWS.</span><span class="sxs-lookup"><span data-stu-id="4159c-141">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-142">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4159c-142">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4159c-p104">Les ID d’élément extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)) utilisent un format différent de celui employé par les services web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.</span><span class="sxs-lookup"><span data-stu-id="4159c-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-145">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-145">Parameters</span></span>

|<span data-ttu-id="4159c-146">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-146">Name</span></span>| <span data-ttu-id="4159c-147">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-147">Type</span></span>| <span data-ttu-id="4159c-148">Description</span><span class="sxs-lookup"><span data-stu-id="4159c-148">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4159c-149">String</span><span class="sxs-lookup"><span data-stu-id="4159c-149">String</span></span>|<span data-ttu-id="4159c-150">ID d’élément mis en forme pour les API REST Outlook</span><span class="sxs-lookup"><span data-stu-id="4159c-150">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="4159c-151">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4159c-151">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="4159c-152">Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.</span><span class="sxs-lookup"><span data-stu-id="4159c-152">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-153">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-153">Requirements</span></span>

|<span data-ttu-id="4159c-154">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-154">Requirement</span></span>| <span data-ttu-id="4159c-155">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-155">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-156">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-156">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-157">1.3</span><span class="sxs-lookup"><span data-stu-id="4159c-157">1.3</span></span>|
|[<span data-ttu-id="4159c-158">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-158">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-159">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4159c-159">Restricted</span></span>|
|[<span data-ttu-id="4159c-160">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-160">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-161">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-161">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4159c-162">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4159c-162">Returns:</span></span>

<span data-ttu-id="4159c-163">Type : String</span><span class="sxs-lookup"><span data-stu-id="4159c-163">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4159c-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-164">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook13officelocalclienttime"></a><span data-ttu-id="4159c-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="4159c-165">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)}</span></span>

<span data-ttu-id="4159c-166">Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.</span><span class="sxs-lookup"><span data-stu-id="4159c-166">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4159c-p105">Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4159c-p105">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4159c-p106">Si l’application de messagerie est en cours d’exécution dans Outlook, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook Web App, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.</span><span class="sxs-lookup"><span data-stu-id="4159c-p106">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-172">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-172">Parameters</span></span>

|<span data-ttu-id="4159c-173">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-173">Name</span></span>| <span data-ttu-id="4159c-174">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-174">Type</span></span>| <span data-ttu-id="4159c-175">Description</span><span class="sxs-lookup"><span data-stu-id="4159c-175">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4159c-176">Date</span><span class="sxs-lookup"><span data-stu-id="4159c-176">Date</span></span>|<span data-ttu-id="4159c-177">Objet Date</span><span class="sxs-lookup"><span data-stu-id="4159c-177">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-178">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-178">Requirements</span></span>

|<span data-ttu-id="4159c-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-179">Requirement</span></span>| <span data-ttu-id="4159c-180">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-181">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-182">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-182">1.0</span></span>|
|[<span data-ttu-id="4159c-183">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-183">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-184">ReadItem</span></span>|
|[<span data-ttu-id="4159c-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-185">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-186">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4159c-187">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4159c-187">Returns:</span></span>

<span data-ttu-id="4159c-188">Type : [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="4159c-188">Type: [LocalClientTime](/javascript/api/outlook_1_3/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="4159c-189">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="4159c-189">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="4159c-190">Convertit un ID d’élément mis en forme pour EWS au format REST.</span><span class="sxs-lookup"><span data-stu-id="4159c-190">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-191">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4159c-191">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4159c-p107">Les ID d’élément récupérés via EWS ou la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](https://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS au format approprié pour REST.</span><span class="sxs-lookup"><span data-stu-id="4159c-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-194">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-194">Parameters</span></span>

|<span data-ttu-id="4159c-195">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-195">Name</span></span>| <span data-ttu-id="4159c-196">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-196">Type</span></span>| <span data-ttu-id="4159c-197">Description</span><span class="sxs-lookup"><span data-stu-id="4159c-197">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4159c-198">String</span><span class="sxs-lookup"><span data-stu-id="4159c-198">String</span></span>|<span data-ttu-id="4159c-199">ID d’élément mis en forme pour les services web Exchange (EWS)</span><span class="sxs-lookup"><span data-stu-id="4159c-199">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="4159c-200">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="4159c-200">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_3/office.mailboxenums.restversion)|<span data-ttu-id="4159c-201">Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.</span><span class="sxs-lookup"><span data-stu-id="4159c-201">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-202">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-202">Requirements</span></span>

|<span data-ttu-id="4159c-203">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-203">Requirement</span></span>| <span data-ttu-id="4159c-204">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-205">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-205">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-206">1.3</span><span class="sxs-lookup"><span data-stu-id="4159c-206">1.3</span></span>|
|[<span data-ttu-id="4159c-207">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-207">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-208">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4159c-208">Restricted</span></span>|
|[<span data-ttu-id="4159c-209">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-210">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-210">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4159c-211">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4159c-211">Returns:</span></span>

<span data-ttu-id="4159c-212">Type : String</span><span class="sxs-lookup"><span data-stu-id="4159c-212">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="4159c-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-213">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4159c-214">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4159c-214">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4159c-215">Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.</span><span class="sxs-lookup"><span data-stu-id="4159c-215">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4159c-216">La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.</span><span class="sxs-lookup"><span data-stu-id="4159c-216">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-217">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-217">Parameters</span></span>

|<span data-ttu-id="4159c-218">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-218">Name</span></span>| <span data-ttu-id="4159c-219">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-219">Type</span></span>| <span data-ttu-id="4159c-220">Description</span><span class="sxs-lookup"><span data-stu-id="4159c-220">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4159c-221">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4159c-221">LocalClientTime</span></span>](/javascript/api/outlook_1_3/office.LocalClientTime)|<span data-ttu-id="4159c-222">Valeur de l’heure locale à convertir.</span><span class="sxs-lookup"><span data-stu-id="4159c-222">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-223">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-223">Requirements</span></span>

|<span data-ttu-id="4159c-224">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-224">Requirement</span></span>| <span data-ttu-id="4159c-225">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-225">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-226">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-226">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-227">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-227">1.0</span></span>|
|[<span data-ttu-id="4159c-228">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-228">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-229">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-229">ReadItem</span></span>|
|[<span data-ttu-id="4159c-230">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-230">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-231">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-231">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4159c-232">Renvoie :</span><span class="sxs-lookup"><span data-stu-id="4159c-232">Returns:</span></span>

<span data-ttu-id="4159c-233">Objet Date avec l’heure exprimée au format UTC.</span><span class="sxs-lookup"><span data-stu-id="4159c-233">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="4159c-234">

<dt>Type</dt>

</span><span class="sxs-lookup"><span data-stu-id="4159c-234">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4159c-235">Date</span><span class="sxs-lookup"><span data-stu-id="4159c-235">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="4159c-236">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4159c-236">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4159c-237">Affiche un rendez-vous de calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4159c-237">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-238">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4159c-238">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4159c-239">La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4159c-239">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4159c-p108">Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.</span><span class="sxs-lookup"><span data-stu-id="4159c-p108">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4159c-242">Dans Outlook Web App, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4159c-242">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4159c-243">Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4159c-243">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-244">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-244">Parameters</span></span>

|<span data-ttu-id="4159c-245">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-245">Name</span></span>| <span data-ttu-id="4159c-246">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-246">Type</span></span>| <span data-ttu-id="4159c-247">object</span><span class="sxs-lookup"><span data-stu-id="4159c-247">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4159c-248">String</span><span class="sxs-lookup"><span data-stu-id="4159c-248">String</span></span>|<span data-ttu-id="4159c-249">Identificateur des services web Exchange pour un rendez-vous du calendrier existant.</span><span class="sxs-lookup"><span data-stu-id="4159c-249">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-250">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-250">Requirements</span></span>

|<span data-ttu-id="4159c-251">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-251">Requirement</span></span>| <span data-ttu-id="4159c-252">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-252">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-253">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-253">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-254">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-254">1.0</span></span>|
|[<span data-ttu-id="4159c-255">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-255">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-256">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-256">ReadItem</span></span>|
|[<span data-ttu-id="4159c-257">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-257">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-258">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4159c-259">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-259">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="4159c-260">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4159c-260">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4159c-261">Affiche un message existant.</span><span class="sxs-lookup"><span data-stu-id="4159c-261">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-262">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4159c-262">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4159c-263">La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="4159c-263">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4159c-264">Dans Outlook Web App, cette méthode ouvre le formulaire indiqué uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4159c-264">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4159c-265">Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.</span><span class="sxs-lookup"><span data-stu-id="4159c-265">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4159c-p109">N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4159c-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-268">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-268">Parameters</span></span>

|<span data-ttu-id="4159c-269">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-269">Name</span></span>| <span data-ttu-id="4159c-270">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-270">Type</span></span>| <span data-ttu-id="4159c-271">object</span><span class="sxs-lookup"><span data-stu-id="4159c-271">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4159c-272">String</span><span class="sxs-lookup"><span data-stu-id="4159c-272">String</span></span>|<span data-ttu-id="4159c-273">Identificateur des services web Exchange pour un message existant.</span><span class="sxs-lookup"><span data-stu-id="4159c-273">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-274">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-274">Requirements</span></span>

|<span data-ttu-id="4159c-275">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-275">Requirement</span></span>| <span data-ttu-id="4159c-276">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-277">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-278">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-278">1.0</span></span>|
|[<span data-ttu-id="4159c-279">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-279">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-280">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-280">ReadItem</span></span>|
|[<span data-ttu-id="4159c-281">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-281">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-282">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-282">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4159c-283">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-283">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4159c-284">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4159c-284">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4159c-285">Affiche un formulaire permettant de créer un rendez-vous du calendrier.</span><span class="sxs-lookup"><span data-stu-id="4159c-285">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-286">Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4159c-286">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4159c-p110">La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.</span><span class="sxs-lookup"><span data-stu-id="4159c-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4159c-p111">Dans Outlook Web App et OWA pour les périphériques, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.</span><span class="sxs-lookup"><span data-stu-id="4159c-p111">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4159c-p112">Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.</span><span class="sxs-lookup"><span data-stu-id="4159c-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4159c-294">Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.</span><span class="sxs-lookup"><span data-stu-id="4159c-294">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-295">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-295">Parameters</span></span>

|<span data-ttu-id="4159c-296">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-296">Name</span></span>| <span data-ttu-id="4159c-297">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-297">Type</span></span>| <span data-ttu-id="4159c-298">object</span><span class="sxs-lookup"><span data-stu-id="4159c-298">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4159c-299">Objet</span><span class="sxs-lookup"><span data-stu-id="4159c-299">Object</span></span> | <span data-ttu-id="4159c-300">Dictionnaire de paramètres décrivant le nouveau rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4159c-300">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4159c-301">Tableau.&lt;Chaîne&gt; &#124; Tableau.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4159c-301">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4159c-p113">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4159c-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4159c-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4159c-304">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_3/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4159c-p114">Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4159c-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4159c-307">Date</span><span class="sxs-lookup"><span data-stu-id="4159c-307">Date</span></span> | <span data-ttu-id="4159c-308">Objet `Date` spécifiant la date et l’heure de début du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4159c-308">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4159c-309">Date</span><span class="sxs-lookup"><span data-stu-id="4159c-309">Date</span></span> | <span data-ttu-id="4159c-310">Objet `Date` spécifiant la date et l’heure de fin du rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="4159c-310">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4159c-311">String</span><span class="sxs-lookup"><span data-stu-id="4159c-311">String</span></span> | <span data-ttu-id="4159c-p115">Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4159c-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4159c-314">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4159c-314">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4159c-p116">Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum.</span><span class="sxs-lookup"><span data-stu-id="4159c-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4159c-317">String</span><span class="sxs-lookup"><span data-stu-id="4159c-317">String</span></span> | <span data-ttu-id="4159c-p117">Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum.</span><span class="sxs-lookup"><span data-stu-id="4159c-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4159c-320">String</span><span class="sxs-lookup"><span data-stu-id="4159c-320">String</span></span> | <span data-ttu-id="4159c-p118">Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko.</span><span class="sxs-lookup"><span data-stu-id="4159c-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4159c-323">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-323">Requirements</span></span>

|<span data-ttu-id="4159c-324">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-324">Requirement</span></span>| <span data-ttu-id="4159c-325">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-326">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-327">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-327">1.0</span></span>|
|[<span data-ttu-id="4159c-328">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-329">ReadItem</span></span>|
|[<span data-ttu-id="4159c-330">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-331">Lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-331">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4159c-332">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-332">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4159c-333">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4159c-333">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4159c-334">Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.</span><span class="sxs-lookup"><span data-stu-id="4159c-334">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4159c-p119">La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.</span><span class="sxs-lookup"><span data-stu-id="4159c-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4159c-p120">Vous pouvez passer le jeton et un identificateur de pièce jointe ou d’élément à un système tiers. Celui-ci utilise le jeton en tant que jeton d’autorisation de support pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services web Exchange (EWS) afin de retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4159c-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4159c-340">Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode de lecture.</span><span class="sxs-lookup"><span data-stu-id="4159c-340">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="4159c-p121">En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.</span><span class="sxs-lookup"><span data-stu-id="4159c-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-343">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-343">Parameters</span></span>

|<span data-ttu-id="4159c-344">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-344">Name</span></span>| <span data-ttu-id="4159c-345">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-345">Type</span></span>| <span data-ttu-id="4159c-346">Attributs</span><span class="sxs-lookup"><span data-stu-id="4159c-346">Attributes</span></span>| <span data-ttu-id="4159c-347">Description</span><span class="sxs-lookup"><span data-stu-id="4159c-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4159c-348">function</span><span class="sxs-lookup"><span data-stu-id="4159c-348">function</span></span>||<span data-ttu-id="4159c-p122">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4159c-p122">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="4159c-351">Objet</span><span class="sxs-lookup"><span data-stu-id="4159c-351">Object</span></span>| <span data-ttu-id="4159c-352">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4159c-352">&lt;optional&gt;</span></span>|<span data-ttu-id="4159c-353">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4159c-353">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-354">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-354">Requirements</span></span>

|<span data-ttu-id="4159c-355">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-355">Requirement</span></span>| <span data-ttu-id="4159c-356">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-356">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-357">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-357">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-358">1.3</span><span class="sxs-lookup"><span data-stu-id="4159c-358">1.3</span></span>|
|[<span data-ttu-id="4159c-359">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-359">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-360">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-360">ReadItem</span></span>|
|[<span data-ttu-id="4159c-361">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-361">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-362">Composition et lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-362">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="4159c-363">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-363">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4159c-364">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4159c-364">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4159c-365">Obtient un jeton qui identifie l’utilisateur et le complément Office.</span><span class="sxs-lookup"><span data-stu-id="4159c-365">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4159c-366">La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="4159c-366">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-367">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-367">Parameters</span></span>

|<span data-ttu-id="4159c-368">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-368">Name</span></span>| <span data-ttu-id="4159c-369">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-369">Type</span></span>| <span data-ttu-id="4159c-370">Attributs</span><span class="sxs-lookup"><span data-stu-id="4159c-370">Attributes</span></span>| <span data-ttu-id="4159c-371">Description</span><span class="sxs-lookup"><span data-stu-id="4159c-371">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4159c-372">fonction</span><span class="sxs-lookup"><span data-stu-id="4159c-372">function</span></span>||<span data-ttu-id="4159c-373">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4159c-373">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4159c-374">Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4159c-374">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="4159c-375">Object</span><span class="sxs-lookup"><span data-stu-id="4159c-375">Object</span></span>| <span data-ttu-id="4159c-376">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4159c-376">&lt;optional&gt;</span></span>|<span data-ttu-id="4159c-377">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4159c-377">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-378">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-378">Requirements</span></span>

|<span data-ttu-id="4159c-379">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-379">Requirement</span></span>| <span data-ttu-id="4159c-380">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-380">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-381">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-381">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-382">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-382">1.0</span></span>|
|[<span data-ttu-id="4159c-383">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-383">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-384">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4159c-384">ReadItem</span></span>|
|[<span data-ttu-id="4159c-385">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-385">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-386">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-386">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4159c-387">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-387">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4159c-388">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4159c-388">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4159c-389">Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4159c-389">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-390">Cette méthode n’est pas prise en charge dans les cas suivants :</span><span class="sxs-lookup"><span data-stu-id="4159c-390">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4159c-391">dans Outlook pour iOS ou Outlook pour Android ;</span><span class="sxs-lookup"><span data-stu-id="4159c-391">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="4159c-392">quand le complément est chargé dans une boîte aux lettres Gmail.</span><span class="sxs-lookup"><span data-stu-id="4159c-392">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4159c-393">Dans ces cas de figure, les compléments doivent [utiliser les API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4159c-393">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4159c-394">La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément.</span><span class="sxs-lookup"><span data-stu-id="4159c-394">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4159c-395">Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="4159c-395">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4159c-396">Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.</span><span class="sxs-lookup"><span data-stu-id="4159c-396">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4159c-397">La demande XML doit spécifier l’encodage UTF-8.</span><span class="sxs-lookup"><span data-stu-id="4159c-397">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4159c-p124">Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="4159c-p124">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4159c-400">L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.</span><span class="sxs-lookup"><span data-stu-id="4159c-400">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4159c-401">Différences entre les versions</span><span class="sxs-lookup"><span data-stu-id="4159c-401">Version differences</span></span>

<span data-ttu-id="4159c-402">Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="4159c-402">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4159c-p125">Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="4159c-p125">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4159c-406">Paramètres</span><span class="sxs-lookup"><span data-stu-id="4159c-406">Parameters</span></span>

|<span data-ttu-id="4159c-407">Nom</span><span class="sxs-lookup"><span data-stu-id="4159c-407">Name</span></span>| <span data-ttu-id="4159c-408">Type</span><span class="sxs-lookup"><span data-stu-id="4159c-408">Type</span></span>| <span data-ttu-id="4159c-409">Attributs</span><span class="sxs-lookup"><span data-stu-id="4159c-409">Attributes</span></span>| <span data-ttu-id="4159c-410">Description</span><span class="sxs-lookup"><span data-stu-id="4159c-410">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4159c-411">String</span><span class="sxs-lookup"><span data-stu-id="4159c-411">String</span></span>||<span data-ttu-id="4159c-412">Demande EWS.</span><span class="sxs-lookup"><span data-stu-id="4159c-412">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4159c-413">function</span><span class="sxs-lookup"><span data-stu-id="4159c-413">function</span></span>||<span data-ttu-id="4159c-414">Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4159c-414">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4159c-415">Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4159c-415">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4159c-416">Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.</span><span class="sxs-lookup"><span data-stu-id="4159c-416">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="4159c-417">Objet</span><span class="sxs-lookup"><span data-stu-id="4159c-417">Object</span></span>| <span data-ttu-id="4159c-418">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4159c-418">&lt;optional&gt;</span></span>|<span data-ttu-id="4159c-419">Données d’état transmises à la méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4159c-419">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4159c-420">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4159c-420">Requirements</span></span>

|<span data-ttu-id="4159c-421">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4159c-421">Requirement</span></span>| <span data-ttu-id="4159c-422">Valeur</span><span class="sxs-lookup"><span data-stu-id="4159c-422">Value</span></span>|
|---|---|
|[<span data-ttu-id="4159c-423">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4159c-423">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4159c-424">1.0</span><span class="sxs-lookup"><span data-stu-id="4159c-424">1.0</span></span>|
|[<span data-ttu-id="4159c-425">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4159c-425">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4159c-426">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4159c-426">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4159c-427">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4159c-427">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4159c-428">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4159c-428">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4159c-429">Exemple</span><span class="sxs-lookup"><span data-stu-id="4159c-429">Example</span></span>

<span data-ttu-id="4159c-430">L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.</span><span class="sxs-lookup"><span data-stu-id="4159c-430">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
