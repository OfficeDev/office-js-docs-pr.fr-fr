---
title: Espace de noms Office – ensemble de conditions requises 1.3
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 9a0f06cbe286f6479ac9244d5ad5bde43ab6b5b6
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457795"
---
# <a name="office"></a><span data-ttu-id="32bc6-102">Office</span><span class="sxs-lookup"><span data-stu-id="32bc6-102">Office</span></span>

<span data-ttu-id="32bc6-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="32bc6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="32bc6-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32bc6-105">Requirements</span></span>

|<span data-ttu-id="32bc6-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32bc6-106">Requirement</span></span>| <span data-ttu-id="32bc6-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="32bc6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="32bc6-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32bc6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32bc6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="32bc6-109">1.0</span></span>|
|[<span data-ttu-id="32bc6-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32bc6-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32bc6-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="32bc6-111">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="32bc6-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="32bc6-112">Namespaces</span></span>

<span data-ttu-id="32bc6-113">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="32bc6-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="32bc6-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="32bc6-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="32bc6-115">Membres</span><span class="sxs-lookup"><span data-stu-id="32bc6-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="32bc6-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="32bc6-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="32bc6-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="32bc6-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="32bc6-118">Type :</span><span class="sxs-lookup"><span data-stu-id="32bc6-118">Type:</span></span>

*   <span data-ttu-id="32bc6-119">String</span><span class="sxs-lookup"><span data-stu-id="32bc6-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32bc6-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="32bc6-120">Properties:</span></span>

|<span data-ttu-id="32bc6-121">Nom</span><span class="sxs-lookup"><span data-stu-id="32bc6-121">Name</span></span>| <span data-ttu-id="32bc6-122">Type</span><span class="sxs-lookup"><span data-stu-id="32bc6-122">Type</span></span>| <span data-ttu-id="32bc6-123">Description</span><span class="sxs-lookup"><span data-stu-id="32bc6-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="32bc6-124">String</span><span class="sxs-lookup"><span data-stu-id="32bc6-124">String</span></span>|<span data-ttu-id="32bc6-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="32bc6-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="32bc6-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="32bc6-126">String</span></span>|<span data-ttu-id="32bc6-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="32bc6-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32bc6-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32bc6-128">Requirements</span></span>

|<span data-ttu-id="32bc6-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32bc6-129">Requirement</span></span>| <span data-ttu-id="32bc6-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="32bc6-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="32bc6-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32bc6-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32bc6-132">1.0</span><span class="sxs-lookup"><span data-stu-id="32bc6-132">1.0</span></span>|
|[<span data-ttu-id="32bc6-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32bc6-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32bc6-134">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="32bc6-134">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="32bc6-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="32bc6-135">CoercionType :String</span></span>

<span data-ttu-id="32bc6-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="32bc6-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="32bc6-137">Type :</span><span class="sxs-lookup"><span data-stu-id="32bc6-137">Type:</span></span>

*   <span data-ttu-id="32bc6-138">String</span><span class="sxs-lookup"><span data-stu-id="32bc6-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32bc6-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="32bc6-139">Properties:</span></span>

|<span data-ttu-id="32bc6-140">Nom</span><span class="sxs-lookup"><span data-stu-id="32bc6-140">Name</span></span>| <span data-ttu-id="32bc6-141">Type</span><span class="sxs-lookup"><span data-stu-id="32bc6-141">Type</span></span>| <span data-ttu-id="32bc6-142">Description</span><span class="sxs-lookup"><span data-stu-id="32bc6-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="32bc6-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="32bc6-143">String</span></span>|<span data-ttu-id="32bc6-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="32bc6-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="32bc6-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="32bc6-145">String</span></span>|<span data-ttu-id="32bc6-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="32bc6-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32bc6-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32bc6-147">Requirements</span></span>

|<span data-ttu-id="32bc6-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32bc6-148">Requirement</span></span>| <span data-ttu-id="32bc6-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="32bc6-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="32bc6-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32bc6-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32bc6-151">1.0</span><span class="sxs-lookup"><span data-stu-id="32bc6-151">1.0</span></span>|
|[<span data-ttu-id="32bc6-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32bc6-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32bc6-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="32bc6-153">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="32bc6-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="32bc6-154">SourceProperty :String</span></span>

<span data-ttu-id="32bc6-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="32bc6-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="32bc6-156">Type :</span><span class="sxs-lookup"><span data-stu-id="32bc6-156">Type:</span></span>

*   <span data-ttu-id="32bc6-157">String</span><span class="sxs-lookup"><span data-stu-id="32bc6-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="32bc6-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="32bc6-158">Properties:</span></span>

|<span data-ttu-id="32bc6-159">Nom</span><span class="sxs-lookup"><span data-stu-id="32bc6-159">Name</span></span>| <span data-ttu-id="32bc6-160">Type</span><span class="sxs-lookup"><span data-stu-id="32bc6-160">Type</span></span>| <span data-ttu-id="32bc6-161">Description</span><span class="sxs-lookup"><span data-stu-id="32bc6-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="32bc6-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="32bc6-162">String</span></span>|<span data-ttu-id="32bc6-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="32bc6-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="32bc6-164">String</span><span class="sxs-lookup"><span data-stu-id="32bc6-164">String</span></span>|<span data-ttu-id="32bc6-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="32bc6-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32bc6-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32bc6-166">Requirements</span></span>

|<span data-ttu-id="32bc6-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32bc6-167">Requirement</span></span>| <span data-ttu-id="32bc6-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="32bc6-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="32bc6-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32bc6-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="32bc6-170">1.0</span><span class="sxs-lookup"><span data-stu-id="32bc6-170">1.0</span></span>|
|[<span data-ttu-id="32bc6-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32bc6-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="32bc6-172">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="32bc6-172">Compose or read</span></span>|