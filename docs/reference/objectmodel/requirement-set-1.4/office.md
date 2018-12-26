---
title: Espace de noms Office – ensemble de conditions requises 1.4
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 7a86c550bd1f40c3db306c518165bc60b8bf0280
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433781"
---
# <a name="office"></a><span data-ttu-id="a9d61-102">Office</span><span class="sxs-lookup"><span data-stu-id="a9d61-102">Office</span></span>

<span data-ttu-id="a9d61-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a9d61-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a9d61-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a9d61-105">Requirements</span></span>

|<span data-ttu-id="a9d61-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a9d61-106">Requirement</span></span>| <span data-ttu-id="a9d61-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="a9d61-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9d61-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a9d61-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9d61-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a9d61-109">1.0</span></span>|
|[<span data-ttu-id="a9d61-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a9d61-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9d61-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a9d61-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="a9d61-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="a9d61-112">Namespaces</span></span>

<span data-ttu-id="a9d61-113">[context](Office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a9d61-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a9d61-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a9d61-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a9d61-115">Membres</span><span class="sxs-lookup"><span data-stu-id="a9d61-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a9d61-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a9d61-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="a9d61-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a9d61-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a9d61-118">Type :</span><span class="sxs-lookup"><span data-stu-id="a9d61-118">Type:</span></span>

*   <span data-ttu-id="a9d61-119">String</span><span class="sxs-lookup"><span data-stu-id="a9d61-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a9d61-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a9d61-120">Properties:</span></span>

|<span data-ttu-id="a9d61-121">Nom</span><span class="sxs-lookup"><span data-stu-id="a9d61-121">Name</span></span>| <span data-ttu-id="a9d61-122">Type</span><span class="sxs-lookup"><span data-stu-id="a9d61-122">Type</span></span>| <span data-ttu-id="a9d61-123">Description</span><span class="sxs-lookup"><span data-stu-id="a9d61-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a9d61-124">String</span><span class="sxs-lookup"><span data-stu-id="a9d61-124">String</span></span>|<span data-ttu-id="a9d61-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="a9d61-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a9d61-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a9d61-126">String</span></span>|<span data-ttu-id="a9d61-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="a9d61-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a9d61-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a9d61-128">Requirements</span></span>

|<span data-ttu-id="a9d61-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a9d61-129">Requirement</span></span>| <span data-ttu-id="a9d61-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="a9d61-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9d61-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a9d61-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9d61-132">1.0</span><span class="sxs-lookup"><span data-stu-id="a9d61-132">1.0</span></span>|
|[<span data-ttu-id="a9d61-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a9d61-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9d61-134">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a9d61-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="a9d61-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a9d61-135">CoercionType :String</span></span>

<span data-ttu-id="a9d61-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a9d61-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a9d61-137">Type :</span><span class="sxs-lookup"><span data-stu-id="a9d61-137">Type:</span></span>

*   <span data-ttu-id="a9d61-138">String</span><span class="sxs-lookup"><span data-stu-id="a9d61-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a9d61-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a9d61-139">Properties:</span></span>

|<span data-ttu-id="a9d61-140">Nom</span><span class="sxs-lookup"><span data-stu-id="a9d61-140">Name</span></span>| <span data-ttu-id="a9d61-141">Type</span><span class="sxs-lookup"><span data-stu-id="a9d61-141">Type</span></span>| <span data-ttu-id="a9d61-142">Description</span><span class="sxs-lookup"><span data-stu-id="a9d61-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a9d61-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a9d61-143">String</span></span>|<span data-ttu-id="a9d61-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="a9d61-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a9d61-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a9d61-145">String</span></span>|<span data-ttu-id="a9d61-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="a9d61-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a9d61-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a9d61-147">Requirements</span></span>

|<span data-ttu-id="a9d61-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a9d61-148">Requirement</span></span>| <span data-ttu-id="a9d61-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="a9d61-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9d61-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a9d61-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9d61-151">1.0</span><span class="sxs-lookup"><span data-stu-id="a9d61-151">1.0</span></span>|
|[<span data-ttu-id="a9d61-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a9d61-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9d61-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a9d61-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="a9d61-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a9d61-154">SourceProperty :String</span></span>

<span data-ttu-id="a9d61-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a9d61-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a9d61-156">Type :</span><span class="sxs-lookup"><span data-stu-id="a9d61-156">Type:</span></span>

*   <span data-ttu-id="a9d61-157">String</span><span class="sxs-lookup"><span data-stu-id="a9d61-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a9d61-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a9d61-158">Properties:</span></span>

|<span data-ttu-id="a9d61-159">Nom</span><span class="sxs-lookup"><span data-stu-id="a9d61-159">Name</span></span>| <span data-ttu-id="a9d61-160">Type</span><span class="sxs-lookup"><span data-stu-id="a9d61-160">Type</span></span>| <span data-ttu-id="a9d61-161">Description</span><span class="sxs-lookup"><span data-stu-id="a9d61-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a9d61-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a9d61-162">String</span></span>|<span data-ttu-id="a9d61-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="a9d61-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a9d61-164">String</span><span class="sxs-lookup"><span data-stu-id="a9d61-164">String</span></span>|<span data-ttu-id="a9d61-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="a9d61-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a9d61-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a9d61-166">Requirements</span></span>

|<span data-ttu-id="a9d61-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a9d61-167">Requirement</span></span>| <span data-ttu-id="a9d61-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="a9d61-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="a9d61-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a9d61-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a9d61-170">1.0</span><span class="sxs-lookup"><span data-stu-id="a9d61-170">1.0</span></span>|
|[<span data-ttu-id="a9d61-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a9d61-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a9d61-172">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a9d61-172">Compose or read</span></span>|