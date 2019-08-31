---
title: Espace de noms Office-ensemble de conditions requises 1,2
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 25322842d82c1db7228a67c268facb6945168222
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696294"
---
# <a name="office"></a><span data-ttu-id="bcff7-102">Office</span><span class="sxs-lookup"><span data-stu-id="bcff7-102">Office</span></span>

<span data-ttu-id="bcff7-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="bcff7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bcff7-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bcff7-105">Requirements</span></span>

|<span data-ttu-id="bcff7-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bcff7-106">Requirement</span></span>| <span data-ttu-id="bcff7-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="bcff7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcff7-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bcff7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bcff7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bcff7-109">1.0</span></span>|
|[<span data-ttu-id="bcff7-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bcff7-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bcff7-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bcff7-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bcff7-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="bcff7-112">Members and methods</span></span>

| <span data-ttu-id="bcff7-113">Membre</span><span class="sxs-lookup"><span data-stu-id="bcff7-113">Member</span></span> | <span data-ttu-id="bcff7-114">Type</span><span class="sxs-lookup"><span data-stu-id="bcff7-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bcff7-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="bcff7-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="bcff7-116">Member</span><span class="sxs-lookup"><span data-stu-id="bcff7-116">Member</span></span> |
| [<span data-ttu-id="bcff7-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="bcff7-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="bcff7-118">Member</span><span class="sxs-lookup"><span data-stu-id="bcff7-118">Member</span></span> |
| [<span data-ttu-id="bcff7-119">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="bcff7-119">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="bcff7-120">Membre</span><span class="sxs-lookup"><span data-stu-id="bcff7-120">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="bcff7-121">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="bcff7-121">Namespaces</span></span>

<span data-ttu-id="bcff7-122">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="bcff7-122">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="bcff7-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): inclut un certain nombre d’énumérations, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="bcff7-123">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.2): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="bcff7-124">Members</span><span class="sxs-lookup"><span data-stu-id="bcff7-124">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="bcff7-125">AsyncResultStatus: chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-125">AsyncResultStatus: String</span></span>

<span data-ttu-id="bcff7-126">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="bcff7-126">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="bcff7-127">Type</span><span class="sxs-lookup"><span data-stu-id="bcff7-127">Type</span></span>

*   <span data-ttu-id="bcff7-128">String</span><span class="sxs-lookup"><span data-stu-id="bcff7-128">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bcff7-129">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="bcff7-129">Properties:</span></span>

|<span data-ttu-id="bcff7-130">Nom</span><span class="sxs-lookup"><span data-stu-id="bcff7-130">Name</span></span>| <span data-ttu-id="bcff7-131">Type</span><span class="sxs-lookup"><span data-stu-id="bcff7-131">Type</span></span>| <span data-ttu-id="bcff7-132">Description</span><span class="sxs-lookup"><span data-stu-id="bcff7-132">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="bcff7-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-133">String</span></span>|<span data-ttu-id="bcff7-134">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="bcff7-134">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="bcff7-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-135">String</span></span>|<span data-ttu-id="bcff7-136">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="bcff7-136">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bcff7-137">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bcff7-137">Requirements</span></span>

|<span data-ttu-id="bcff7-138">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bcff7-138">Requirement</span></span>| <span data-ttu-id="bcff7-139">Valeur</span><span class="sxs-lookup"><span data-stu-id="bcff7-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcff7-140">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bcff7-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bcff7-141">1.0</span><span class="sxs-lookup"><span data-stu-id="bcff7-141">1.0</span></span>|
|[<span data-ttu-id="bcff7-142">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bcff7-142">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bcff7-143">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bcff7-143">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="bcff7-144">CoercionType: chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-144">CoercionType: String</span></span>

<span data-ttu-id="bcff7-145">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="bcff7-145">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bcff7-146">Type</span><span class="sxs-lookup"><span data-stu-id="bcff7-146">Type</span></span>

*   <span data-ttu-id="bcff7-147">String</span><span class="sxs-lookup"><span data-stu-id="bcff7-147">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bcff7-148">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="bcff7-148">Properties:</span></span>

|<span data-ttu-id="bcff7-149">Nom</span><span class="sxs-lookup"><span data-stu-id="bcff7-149">Name</span></span>| <span data-ttu-id="bcff7-150">Type</span><span class="sxs-lookup"><span data-stu-id="bcff7-150">Type</span></span>| <span data-ttu-id="bcff7-151">Description</span><span class="sxs-lookup"><span data-stu-id="bcff7-151">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="bcff7-152">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-152">String</span></span>|<span data-ttu-id="bcff7-153">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="bcff7-153">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="bcff7-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-154">String</span></span>|<span data-ttu-id="bcff7-155">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="bcff7-155">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bcff7-156">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bcff7-156">Requirements</span></span>

|<span data-ttu-id="bcff7-157">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bcff7-157">Requirement</span></span>| <span data-ttu-id="bcff7-158">Valeur</span><span class="sxs-lookup"><span data-stu-id="bcff7-158">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcff7-159">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bcff7-159">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bcff7-160">1.0</span><span class="sxs-lookup"><span data-stu-id="bcff7-160">1.0</span></span>|
|[<span data-ttu-id="bcff7-161">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bcff7-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bcff7-162">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bcff7-162">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="bcff7-163">SourceProperty: chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-163">SourceProperty: String</span></span>

<span data-ttu-id="bcff7-164">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="bcff7-164">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bcff7-165">Type</span><span class="sxs-lookup"><span data-stu-id="bcff7-165">Type</span></span>

*   <span data-ttu-id="bcff7-166">String</span><span class="sxs-lookup"><span data-stu-id="bcff7-166">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bcff7-167">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="bcff7-167">Properties:</span></span>

|<span data-ttu-id="bcff7-168">Nom</span><span class="sxs-lookup"><span data-stu-id="bcff7-168">Name</span></span>| <span data-ttu-id="bcff7-169">Type</span><span class="sxs-lookup"><span data-stu-id="bcff7-169">Type</span></span>| <span data-ttu-id="bcff7-170">Description</span><span class="sxs-lookup"><span data-stu-id="bcff7-170">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="bcff7-171">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-171">String</span></span>|<span data-ttu-id="bcff7-172">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="bcff7-172">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="bcff7-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bcff7-173">String</span></span>|<span data-ttu-id="bcff7-174">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="bcff7-174">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bcff7-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bcff7-175">Requirements</span></span>

|<span data-ttu-id="bcff7-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bcff7-176">Requirement</span></span>| <span data-ttu-id="bcff7-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="bcff7-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="bcff7-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bcff7-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bcff7-179">1.0</span><span class="sxs-lookup"><span data-stu-id="bcff7-179">1.0</span></span>|
|[<span data-ttu-id="bcff7-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bcff7-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bcff7-181">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bcff7-181">Compose or Read</span></span>|
