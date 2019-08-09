---
title: Espace de noms Office-ensemble de conditions requises 1,4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: e6c4614af74a665805c400c407e4a7785efe9f96
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268648"
---
# <a name="office"></a><span data-ttu-id="c1582-102">Office</span><span class="sxs-lookup"><span data-stu-id="c1582-102">Office</span></span>

<span data-ttu-id="c1582-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c1582-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1582-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1582-105">Requirements</span></span>

|<span data-ttu-id="c1582-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1582-106">Requirement</span></span>| <span data-ttu-id="c1582-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1582-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1582-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1582-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1582-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c1582-109">1.0</span></span>|
|[<span data-ttu-id="c1582-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1582-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1582-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1582-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c1582-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c1582-112">Members and methods</span></span>

| <span data-ttu-id="c1582-113">Membre</span><span class="sxs-lookup"><span data-stu-id="c1582-113">Member</span></span> | <span data-ttu-id="c1582-114">Type</span><span class="sxs-lookup"><span data-stu-id="c1582-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c1582-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c1582-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c1582-116">Member</span><span class="sxs-lookup"><span data-stu-id="c1582-116">Member</span></span> |
| [<span data-ttu-id="c1582-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c1582-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c1582-118">Member</span><span class="sxs-lookup"><span data-stu-id="c1582-118">Member</span></span> |
| [<span data-ttu-id="c1582-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c1582-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c1582-120">Membre</span><span class="sxs-lookup"><span data-stu-id="c1582-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c1582-121">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c1582-121">Namespaces</span></span>

<span data-ttu-id="c1582-122">[context](Office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="c1582-122">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c1582-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c1582-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c1582-124">Membres</span><span class="sxs-lookup"><span data-stu-id="c1582-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="c1582-125">AsyncResultStatus: chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="c1582-126">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c1582-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c1582-127">Type</span><span class="sxs-lookup"><span data-stu-id="c1582-127">Type</span></span>

*   <span data-ttu-id="c1582-128">String</span><span class="sxs-lookup"><span data-stu-id="c1582-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1582-129">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c1582-129">Properties:</span></span>

|<span data-ttu-id="c1582-130">Nom</span><span class="sxs-lookup"><span data-stu-id="c1582-130">Name</span></span>| <span data-ttu-id="c1582-131">Type</span><span class="sxs-lookup"><span data-stu-id="c1582-131">Type</span></span>| <span data-ttu-id="c1582-132">Description</span><span class="sxs-lookup"><span data-stu-id="c1582-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c1582-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-133">String</span></span>|<span data-ttu-id="c1582-134">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="c1582-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c1582-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-135">String</span></span>|<span data-ttu-id="c1582-136">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="c1582-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1582-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1582-137">Requirements</span></span>

|<span data-ttu-id="c1582-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1582-138">Requirement</span></span>| <span data-ttu-id="c1582-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1582-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1582-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1582-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1582-141">1.0</span><span class="sxs-lookup"><span data-stu-id="c1582-141">1.0</span></span>|
|[<span data-ttu-id="c1582-142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1582-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1582-143">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1582-143">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="c1582-144">CoercionType: chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-144">CoercionType: String</span></span>

<span data-ttu-id="c1582-145">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c1582-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c1582-146">Type</span><span class="sxs-lookup"><span data-stu-id="c1582-146">Type</span></span>

*   <span data-ttu-id="c1582-147">String</span><span class="sxs-lookup"><span data-stu-id="c1582-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1582-148">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c1582-148">Properties:</span></span>

|<span data-ttu-id="c1582-149">Nom</span><span class="sxs-lookup"><span data-stu-id="c1582-149">Name</span></span>| <span data-ttu-id="c1582-150">Type</span><span class="sxs-lookup"><span data-stu-id="c1582-150">Type</span></span>| <span data-ttu-id="c1582-151">Description</span><span class="sxs-lookup"><span data-stu-id="c1582-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c1582-152">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-152">String</span></span>|<span data-ttu-id="c1582-153">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="c1582-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c1582-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-154">String</span></span>|<span data-ttu-id="c1582-155">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="c1582-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1582-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1582-156">Requirements</span></span>

|<span data-ttu-id="c1582-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1582-157">Requirement</span></span>| <span data-ttu-id="c1582-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1582-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1582-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1582-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1582-160">1.0</span><span class="sxs-lookup"><span data-stu-id="c1582-160">1.0</span></span>|
|[<span data-ttu-id="c1582-161">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1582-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1582-162">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1582-162">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="c1582-163">SourceProperty: chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-163">SourceProperty: String</span></span>

<span data-ttu-id="c1582-164">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c1582-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c1582-165">Type</span><span class="sxs-lookup"><span data-stu-id="c1582-165">Type</span></span>

*   <span data-ttu-id="c1582-166">String</span><span class="sxs-lookup"><span data-stu-id="c1582-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1582-167">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c1582-167">Properties:</span></span>

|<span data-ttu-id="c1582-168">Nom</span><span class="sxs-lookup"><span data-stu-id="c1582-168">Name</span></span>| <span data-ttu-id="c1582-169">Type</span><span class="sxs-lookup"><span data-stu-id="c1582-169">Type</span></span>| <span data-ttu-id="c1582-170">Description</span><span class="sxs-lookup"><span data-stu-id="c1582-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c1582-171">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-171">String</span></span>|<span data-ttu-id="c1582-172">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="c1582-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c1582-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1582-173">String</span></span>|<span data-ttu-id="c1582-174">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="c1582-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1582-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1582-175">Requirements</span></span>

|<span data-ttu-id="c1582-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1582-176">Requirement</span></span>| <span data-ttu-id="c1582-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1582-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1582-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1582-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1582-179">1.0</span><span class="sxs-lookup"><span data-stu-id="c1582-179">1.0</span></span>|
|[<span data-ttu-id="c1582-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1582-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c1582-181">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1582-181">Compose or Read</span></span>|
