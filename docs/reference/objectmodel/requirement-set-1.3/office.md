---
title: Espace de noms Office-ensemble de conditions requises 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 0b22574693fb129be6a08a89b58beceb746fa283
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268396"
---
# <a name="office"></a><span data-ttu-id="12a8f-102">Office</span><span class="sxs-lookup"><span data-stu-id="12a8f-102">Office</span></span>

<span data-ttu-id="12a8f-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="12a8f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="12a8f-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12a8f-105">Requirements</span></span>

|<span data-ttu-id="12a8f-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12a8f-106">Requirement</span></span>| <span data-ttu-id="12a8f-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="12a8f-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="12a8f-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12a8f-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12a8f-109">1.0</span><span class="sxs-lookup"><span data-stu-id="12a8f-109">1.0</span></span>|
|[<span data-ttu-id="12a8f-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12a8f-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12a8f-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12a8f-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="12a8f-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="12a8f-112">Members and methods</span></span>

| <span data-ttu-id="12a8f-113">Membre</span><span class="sxs-lookup"><span data-stu-id="12a8f-113">Member</span></span> | <span data-ttu-id="12a8f-114">Type</span><span class="sxs-lookup"><span data-stu-id="12a8f-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="12a8f-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="12a8f-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="12a8f-116">Member</span><span class="sxs-lookup"><span data-stu-id="12a8f-116">Member</span></span> |
| [<span data-ttu-id="12a8f-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="12a8f-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="12a8f-118">Member</span><span class="sxs-lookup"><span data-stu-id="12a8f-118">Member</span></span> |
| [<span data-ttu-id="12a8f-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="12a8f-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="12a8f-120">Membre</span><span class="sxs-lookup"><span data-stu-id="12a8f-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="12a8f-121">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="12a8f-121">Namespaces</span></span>

<span data-ttu-id="12a8f-122">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="12a8f-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="12a8f-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="12a8f-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.3): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="12a8f-124">Membres</span><span class="sxs-lookup"><span data-stu-id="12a8f-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="12a8f-125">AsyncResultStatus: chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="12a8f-126">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="12a8f-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="12a8f-127">Type</span><span class="sxs-lookup"><span data-stu-id="12a8f-127">Type</span></span>

*   <span data-ttu-id="12a8f-128">String</span><span class="sxs-lookup"><span data-stu-id="12a8f-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="12a8f-129">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="12a8f-129">Properties:</span></span>

|<span data-ttu-id="12a8f-130">Nom</span><span class="sxs-lookup"><span data-stu-id="12a8f-130">Name</span></span>| <span data-ttu-id="12a8f-131">Type</span><span class="sxs-lookup"><span data-stu-id="12a8f-131">Type</span></span>| <span data-ttu-id="12a8f-132">Description</span><span class="sxs-lookup"><span data-stu-id="12a8f-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="12a8f-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-133">String</span></span>|<span data-ttu-id="12a8f-134">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="12a8f-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="12a8f-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-135">String</span></span>|<span data-ttu-id="12a8f-136">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="12a8f-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12a8f-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12a8f-137">Requirements</span></span>

|<span data-ttu-id="12a8f-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12a8f-138">Requirement</span></span>| <span data-ttu-id="12a8f-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="12a8f-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="12a8f-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12a8f-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12a8f-141">1.0</span><span class="sxs-lookup"><span data-stu-id="12a8f-141">1.0</span></span>|
|[<span data-ttu-id="12a8f-142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12a8f-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12a8f-143">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12a8f-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="12a8f-144">CoercionType: chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-144">CoercionType: String</span></span>

<span data-ttu-id="12a8f-145">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="12a8f-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="12a8f-146">Type</span><span class="sxs-lookup"><span data-stu-id="12a8f-146">Type</span></span>

*   <span data-ttu-id="12a8f-147">String</span><span class="sxs-lookup"><span data-stu-id="12a8f-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="12a8f-148">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="12a8f-148">Properties:</span></span>

|<span data-ttu-id="12a8f-149">Nom</span><span class="sxs-lookup"><span data-stu-id="12a8f-149">Name</span></span>| <span data-ttu-id="12a8f-150">Type</span><span class="sxs-lookup"><span data-stu-id="12a8f-150">Type</span></span>| <span data-ttu-id="12a8f-151">Description</span><span class="sxs-lookup"><span data-stu-id="12a8f-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="12a8f-152">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-152">String</span></span>|<span data-ttu-id="12a8f-153">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="12a8f-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="12a8f-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-154">String</span></span>|<span data-ttu-id="12a8f-155">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="12a8f-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12a8f-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12a8f-156">Requirements</span></span>

|<span data-ttu-id="12a8f-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12a8f-157">Requirement</span></span>| <span data-ttu-id="12a8f-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="12a8f-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="12a8f-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12a8f-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12a8f-160">1.0</span><span class="sxs-lookup"><span data-stu-id="12a8f-160">1.0</span></span>|
|[<span data-ttu-id="12a8f-161">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12a8f-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12a8f-162">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12a8f-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="12a8f-163">SourceProperty: chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-163">SourceProperty: String</span></span>

<span data-ttu-id="12a8f-164">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="12a8f-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="12a8f-165">Type</span><span class="sxs-lookup"><span data-stu-id="12a8f-165">Type</span></span>

*   <span data-ttu-id="12a8f-166">String</span><span class="sxs-lookup"><span data-stu-id="12a8f-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="12a8f-167">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="12a8f-167">Properties:</span></span>

|<span data-ttu-id="12a8f-168">Nom</span><span class="sxs-lookup"><span data-stu-id="12a8f-168">Name</span></span>| <span data-ttu-id="12a8f-169">Type</span><span class="sxs-lookup"><span data-stu-id="12a8f-169">Type</span></span>| <span data-ttu-id="12a8f-170">Description</span><span class="sxs-lookup"><span data-stu-id="12a8f-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="12a8f-171">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-171">String</span></span>|<span data-ttu-id="12a8f-172">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="12a8f-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="12a8f-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="12a8f-173">String</span></span>|<span data-ttu-id="12a8f-174">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="12a8f-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="12a8f-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="12a8f-175">Requirements</span></span>

|<span data-ttu-id="12a8f-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="12a8f-176">Requirement</span></span>| <span data-ttu-id="12a8f-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="12a8f-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="12a8f-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="12a8f-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12a8f-179">1.0</span><span class="sxs-lookup"><span data-stu-id="12a8f-179">1.0</span></span>|
|[<span data-ttu-id="12a8f-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="12a8f-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="12a8f-181">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="12a8f-181">Compose or Read</span></span>|
