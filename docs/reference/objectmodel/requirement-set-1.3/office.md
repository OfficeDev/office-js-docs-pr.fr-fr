---
title: Espace de noms Office – ensemble de conditions requises 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: ad08895719d809436216d2f0bb455260dbca3b1e
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067887"
---
# <a name="office"></a><span data-ttu-id="a4432-102">Office</span><span class="sxs-lookup"><span data-stu-id="a4432-102">Office</span></span>

<span data-ttu-id="a4432-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a4432-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a4432-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a4432-105">Requirements</span></span>

|<span data-ttu-id="a4432-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a4432-106">Requirement</span></span>| <span data-ttu-id="a4432-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="a4432-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4432-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a4432-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4432-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a4432-109">1.0</span></span>|
|[<span data-ttu-id="a4432-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a4432-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4432-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a4432-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="a4432-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="a4432-112">Namespaces</span></span>

<span data-ttu-id="a4432-113">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a4432-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a4432-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="a4432-114">[MailboxEnums](/javascript/api/outlook_1_3/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="a4432-115">Membres</span><span class="sxs-lookup"><span data-stu-id="a4432-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="a4432-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="a4432-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="a4432-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a4432-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a4432-118">Type</span><span class="sxs-lookup"><span data-stu-id="a4432-118">Type</span></span>

*   <span data-ttu-id="a4432-119">String</span><span class="sxs-lookup"><span data-stu-id="a4432-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a4432-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a4432-120">Properties:</span></span>

|<span data-ttu-id="a4432-121">Nom</span><span class="sxs-lookup"><span data-stu-id="a4432-121">Name</span></span>| <span data-ttu-id="a4432-122">Type</span><span class="sxs-lookup"><span data-stu-id="a4432-122">Type</span></span>| <span data-ttu-id="a4432-123">Description</span><span class="sxs-lookup"><span data-stu-id="a4432-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a4432-124">String</span><span class="sxs-lookup"><span data-stu-id="a4432-124">String</span></span>|<span data-ttu-id="a4432-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="a4432-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a4432-126">String</span><span class="sxs-lookup"><span data-stu-id="a4432-126">String</span></span>|<span data-ttu-id="a4432-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="a4432-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a4432-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a4432-128">Requirements</span></span>

|<span data-ttu-id="a4432-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a4432-129">Requirement</span></span>| <span data-ttu-id="a4432-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="a4432-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4432-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a4432-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4432-132">1.0</span><span class="sxs-lookup"><span data-stu-id="a4432-132">1.0</span></span>|
|[<span data-ttu-id="a4432-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a4432-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4432-134">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a4432-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="a4432-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="a4432-135">CoercionType :String</span></span>

<span data-ttu-id="a4432-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a4432-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a4432-137">Type</span><span class="sxs-lookup"><span data-stu-id="a4432-137">Type</span></span>

*   <span data-ttu-id="a4432-138">String</span><span class="sxs-lookup"><span data-stu-id="a4432-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a4432-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a4432-139">Properties:</span></span>

|<span data-ttu-id="a4432-140">Nom</span><span class="sxs-lookup"><span data-stu-id="a4432-140">Name</span></span>| <span data-ttu-id="a4432-141">Type</span><span class="sxs-lookup"><span data-stu-id="a4432-141">Type</span></span>| <span data-ttu-id="a4432-142">Description</span><span class="sxs-lookup"><span data-stu-id="a4432-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a4432-143">String</span><span class="sxs-lookup"><span data-stu-id="a4432-143">String</span></span>|<span data-ttu-id="a4432-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="a4432-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a4432-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a4432-145">String</span></span>|<span data-ttu-id="a4432-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="a4432-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a4432-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a4432-147">Requirements</span></span>

|<span data-ttu-id="a4432-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a4432-148">Requirement</span></span>| <span data-ttu-id="a4432-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="a4432-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4432-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a4432-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4432-151">1.0</span><span class="sxs-lookup"><span data-stu-id="a4432-151">1.0</span></span>|
|[<span data-ttu-id="a4432-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a4432-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4432-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a4432-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="a4432-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="a4432-154">SourceProperty :String</span></span>

<span data-ttu-id="a4432-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a4432-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a4432-156">Type</span><span class="sxs-lookup"><span data-stu-id="a4432-156">Type</span></span>

*   <span data-ttu-id="a4432-157">String</span><span class="sxs-lookup"><span data-stu-id="a4432-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a4432-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a4432-158">Properties:</span></span>

|<span data-ttu-id="a4432-159">Nom</span><span class="sxs-lookup"><span data-stu-id="a4432-159">Name</span></span>| <span data-ttu-id="a4432-160">Type</span><span class="sxs-lookup"><span data-stu-id="a4432-160">Type</span></span>| <span data-ttu-id="a4432-161">Description</span><span class="sxs-lookup"><span data-stu-id="a4432-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a4432-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a4432-162">String</span></span>|<span data-ttu-id="a4432-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="a4432-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a4432-164">String</span><span class="sxs-lookup"><span data-stu-id="a4432-164">String</span></span>|<span data-ttu-id="a4432-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="a4432-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a4432-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a4432-166">Requirements</span></span>

|<span data-ttu-id="a4432-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a4432-167">Requirement</span></span>| <span data-ttu-id="a4432-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="a4432-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="a4432-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a4432-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a4432-170">1.0</span><span class="sxs-lookup"><span data-stu-id="a4432-170">1.0</span></span>|
|[<span data-ttu-id="a4432-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a4432-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a4432-172">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="a4432-172">Compose or Read</span></span>|
