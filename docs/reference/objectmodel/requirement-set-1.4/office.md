---
title: Espace de noms Office-ensemble de conditions requises 1,4
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: c60195ddfc42d962427127bf601bca3d41797566
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30872108"
---
# <a name="office"></a><span data-ttu-id="e2373-102">Office</span><span class="sxs-lookup"><span data-stu-id="e2373-102">Office</span></span>

<span data-ttu-id="e2373-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e2373-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e2373-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e2373-105">Requirements</span></span>

|<span data-ttu-id="e2373-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e2373-106">Requirement</span></span>| <span data-ttu-id="e2373-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="e2373-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2373-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e2373-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2373-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e2373-109">1.0</span></span>|
|[<span data-ttu-id="e2373-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e2373-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2373-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e2373-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="e2373-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="e2373-112">Namespaces</span></span>

<span data-ttu-id="e2373-113">[context](Office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="e2373-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="e2373-114">[MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="e2373-114">[MailboxEnums](/javascript/api/outlook_1_4/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="e2373-115">Membres</span><span class="sxs-lookup"><span data-stu-id="e2373-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="e2373-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="e2373-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="e2373-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="e2373-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e2373-118">Type</span><span class="sxs-lookup"><span data-stu-id="e2373-118">Type</span></span>

*   <span data-ttu-id="e2373-119">String</span><span class="sxs-lookup"><span data-stu-id="e2373-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e2373-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="e2373-120">Properties:</span></span>

|<span data-ttu-id="e2373-121">Nom</span><span class="sxs-lookup"><span data-stu-id="e2373-121">Name</span></span>| <span data-ttu-id="e2373-122">Type</span><span class="sxs-lookup"><span data-stu-id="e2373-122">Type</span></span>| <span data-ttu-id="e2373-123">Description</span><span class="sxs-lookup"><span data-stu-id="e2373-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e2373-124">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e2373-124">String</span></span>|<span data-ttu-id="e2373-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="e2373-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e2373-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e2373-126">String</span></span>|<span data-ttu-id="e2373-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="e2373-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e2373-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e2373-128">Requirements</span></span>

|<span data-ttu-id="e2373-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e2373-129">Requirement</span></span>| <span data-ttu-id="e2373-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="e2373-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2373-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e2373-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2373-132">1.0</span><span class="sxs-lookup"><span data-stu-id="e2373-132">1.0</span></span>|
|[<span data-ttu-id="e2373-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e2373-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2373-134">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e2373-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="e2373-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="e2373-135">CoercionType :String</span></span>

<span data-ttu-id="e2373-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e2373-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e2373-137">Type</span><span class="sxs-lookup"><span data-stu-id="e2373-137">Type</span></span>

*   <span data-ttu-id="e2373-138">String</span><span class="sxs-lookup"><span data-stu-id="e2373-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e2373-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="e2373-139">Properties:</span></span>

|<span data-ttu-id="e2373-140">Nom</span><span class="sxs-lookup"><span data-stu-id="e2373-140">Name</span></span>| <span data-ttu-id="e2373-141">Type</span><span class="sxs-lookup"><span data-stu-id="e2373-141">Type</span></span>| <span data-ttu-id="e2373-142">Description</span><span class="sxs-lookup"><span data-stu-id="e2373-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e2373-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e2373-143">String</span></span>|<span data-ttu-id="e2373-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="e2373-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e2373-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e2373-145">String</span></span>|<span data-ttu-id="e2373-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="e2373-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e2373-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e2373-147">Requirements</span></span>

|<span data-ttu-id="e2373-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e2373-148">Requirement</span></span>| <span data-ttu-id="e2373-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="e2373-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2373-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e2373-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2373-151">1.0</span><span class="sxs-lookup"><span data-stu-id="e2373-151">1.0</span></span>|
|[<span data-ttu-id="e2373-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e2373-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2373-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e2373-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="e2373-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="e2373-154">SourceProperty :String</span></span>

<span data-ttu-id="e2373-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e2373-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e2373-156">Type</span><span class="sxs-lookup"><span data-stu-id="e2373-156">Type</span></span>

*   <span data-ttu-id="e2373-157">String</span><span class="sxs-lookup"><span data-stu-id="e2373-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e2373-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="e2373-158">Properties:</span></span>

|<span data-ttu-id="e2373-159">Nom</span><span class="sxs-lookup"><span data-stu-id="e2373-159">Name</span></span>| <span data-ttu-id="e2373-160">Type</span><span class="sxs-lookup"><span data-stu-id="e2373-160">Type</span></span>| <span data-ttu-id="e2373-161">Description</span><span class="sxs-lookup"><span data-stu-id="e2373-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e2373-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e2373-162">String</span></span>|<span data-ttu-id="e2373-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="e2373-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e2373-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e2373-164">String</span></span>|<span data-ttu-id="e2373-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="e2373-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e2373-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e2373-166">Requirements</span></span>

|<span data-ttu-id="e2373-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e2373-167">Requirement</span></span>| <span data-ttu-id="e2373-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="e2373-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="e2373-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e2373-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e2373-170">1.0</span><span class="sxs-lookup"><span data-stu-id="e2373-170">1.0</span></span>|
|[<span data-ttu-id="e2373-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e2373-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e2373-172">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e2373-172">Compose or Read</span></span>|
