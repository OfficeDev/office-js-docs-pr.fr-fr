---
title: Espace de noms Office – ensemble de conditions requises 1.3
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: c269f21b98e7f87d6f064f6c8ea0c439916f7caf
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432745"
---
# <a name="office"></a><span data-ttu-id="767bb-102">Office</span><span class="sxs-lookup"><span data-stu-id="767bb-102">Office</span></span>

<span data-ttu-id="767bb-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="767bb-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="767bb-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="767bb-105">Requirements</span></span>

|<span data-ttu-id="767bb-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="767bb-106">Requirement</span></span>| <span data-ttu-id="767bb-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="767bb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="767bb-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="767bb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="767bb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="767bb-109">1.0</span></span>|
|[<span data-ttu-id="767bb-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="767bb-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="767bb-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="767bb-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="767bb-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="767bb-112">Namespaces</span></span>

<span data-ttu-id="767bb-113">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="767bb-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="767bb-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="767bb-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="767bb-115">Membres</span><span class="sxs-lookup"><span data-stu-id="767bb-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="767bb-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="767bb-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="767bb-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="767bb-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="767bb-118">Type :</span><span class="sxs-lookup"><span data-stu-id="767bb-118">Type:</span></span>

*   <span data-ttu-id="767bb-119">String</span><span class="sxs-lookup"><span data-stu-id="767bb-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="767bb-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="767bb-120">Properties:</span></span>

|<span data-ttu-id="767bb-121">Nom</span><span class="sxs-lookup"><span data-stu-id="767bb-121">Name</span></span>| <span data-ttu-id="767bb-122">Type</span><span class="sxs-lookup"><span data-stu-id="767bb-122">Type</span></span>| <span data-ttu-id="767bb-123">Description</span><span class="sxs-lookup"><span data-stu-id="767bb-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="767bb-124">String</span><span class="sxs-lookup"><span data-stu-id="767bb-124">String</span></span>|<span data-ttu-id="767bb-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="767bb-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="767bb-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="767bb-126">String</span></span>|<span data-ttu-id="767bb-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="767bb-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="767bb-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="767bb-128">Requirements</span></span>

|<span data-ttu-id="767bb-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="767bb-129">Requirement</span></span>| <span data-ttu-id="767bb-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="767bb-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="767bb-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="767bb-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="767bb-132">1.0</span><span class="sxs-lookup"><span data-stu-id="767bb-132">1.0</span></span>|
|[<span data-ttu-id="767bb-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="767bb-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="767bb-134">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="767bb-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="767bb-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="767bb-135">CoercionType :String</span></span>

<span data-ttu-id="767bb-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="767bb-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="767bb-137">Type :</span><span class="sxs-lookup"><span data-stu-id="767bb-137">Type:</span></span>

*   <span data-ttu-id="767bb-138">String</span><span class="sxs-lookup"><span data-stu-id="767bb-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="767bb-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="767bb-139">Properties:</span></span>

|<span data-ttu-id="767bb-140">Nom</span><span class="sxs-lookup"><span data-stu-id="767bb-140">Name</span></span>| <span data-ttu-id="767bb-141">Type</span><span class="sxs-lookup"><span data-stu-id="767bb-141">Type</span></span>| <span data-ttu-id="767bb-142">Description</span><span class="sxs-lookup"><span data-stu-id="767bb-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="767bb-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="767bb-143">String</span></span>|<span data-ttu-id="767bb-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="767bb-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="767bb-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="767bb-145">String</span></span>|<span data-ttu-id="767bb-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="767bb-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="767bb-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="767bb-147">Requirements</span></span>

|<span data-ttu-id="767bb-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="767bb-148">Requirement</span></span>| <span data-ttu-id="767bb-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="767bb-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="767bb-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="767bb-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="767bb-151">1.0</span><span class="sxs-lookup"><span data-stu-id="767bb-151">1.0</span></span>|
|[<span data-ttu-id="767bb-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="767bb-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="767bb-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="767bb-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="767bb-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="767bb-154">SourceProperty :String</span></span>

<span data-ttu-id="767bb-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="767bb-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="767bb-156">Type :</span><span class="sxs-lookup"><span data-stu-id="767bb-156">Type:</span></span>

*   <span data-ttu-id="767bb-157">String</span><span class="sxs-lookup"><span data-stu-id="767bb-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="767bb-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="767bb-158">Properties:</span></span>

|<span data-ttu-id="767bb-159">Nom</span><span class="sxs-lookup"><span data-stu-id="767bb-159">Name</span></span>| <span data-ttu-id="767bb-160">Type</span><span class="sxs-lookup"><span data-stu-id="767bb-160">Type</span></span>| <span data-ttu-id="767bb-161">Description</span><span class="sxs-lookup"><span data-stu-id="767bb-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="767bb-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="767bb-162">String</span></span>|<span data-ttu-id="767bb-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="767bb-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="767bb-164">String</span><span class="sxs-lookup"><span data-stu-id="767bb-164">String</span></span>|<span data-ttu-id="767bb-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="767bb-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="767bb-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="767bb-166">Requirements</span></span>

|<span data-ttu-id="767bb-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="767bb-167">Requirement</span></span>| <span data-ttu-id="767bb-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="767bb-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="767bb-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="767bb-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="767bb-170">1.0</span><span class="sxs-lookup"><span data-stu-id="767bb-170">1.0</span></span>|
|[<span data-ttu-id="767bb-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="767bb-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="767bb-172">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="767bb-172">Compose or read</span></span>|