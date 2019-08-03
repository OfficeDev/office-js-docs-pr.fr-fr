---
title: Espace de noms Office-ensemble de conditions requises 1,4
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 2617ba3c80f44b1cddab568f94044a95a3061065
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064681"
---
# <a name="office"></a><span data-ttu-id="eb9c0-102">Office</span><span class="sxs-lookup"><span data-stu-id="eb9c0-102">Office</span></span>

<span data-ttu-id="eb9c0-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="eb9c0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb9c0-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eb9c0-105">Requirements</span></span>

|<span data-ttu-id="eb9c0-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eb9c0-106">Requirement</span></span>| <span data-ttu-id="eb9c0-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="eb9c0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb9c0-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eb9c0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb9c0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="eb9c0-109">1.0</span></span>|
|[<span data-ttu-id="eb9c0-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eb9c0-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb9c0-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="eb9c0-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="eb9c0-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="eb9c0-112">Namespaces</span></span>

<span data-ttu-id="eb9c0-113">[context](Office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-113">[context](Office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="eb9c0-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-114">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.4): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="eb9c0-115">Membres</span><span class="sxs-lookup"><span data-stu-id="eb9c0-115">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="eb9c0-116">AsyncResultStatus: chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-116">AsyncResultStatus: String</span></span>

<span data-ttu-id="eb9c0-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="eb9c0-118">Type</span><span class="sxs-lookup"><span data-stu-id="eb9c0-118">Type</span></span>

*   <span data-ttu-id="eb9c0-119">String</span><span class="sxs-lookup"><span data-stu-id="eb9c0-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eb9c0-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="eb9c0-120">Properties:</span></span>

|<span data-ttu-id="eb9c0-121">Nom</span><span class="sxs-lookup"><span data-stu-id="eb9c0-121">Name</span></span>| <span data-ttu-id="eb9c0-122">Type</span><span class="sxs-lookup"><span data-stu-id="eb9c0-122">Type</span></span>| <span data-ttu-id="eb9c0-123">Description</span><span class="sxs-lookup"><span data-stu-id="eb9c0-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="eb9c0-124">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-124">String</span></span>|<span data-ttu-id="eb9c0-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="eb9c0-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-126">String</span></span>|<span data-ttu-id="eb9c0-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb9c0-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eb9c0-128">Requirements</span></span>

|<span data-ttu-id="eb9c0-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eb9c0-129">Requirement</span></span>| <span data-ttu-id="eb9c0-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="eb9c0-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb9c0-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eb9c0-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb9c0-132">1.0</span><span class="sxs-lookup"><span data-stu-id="eb9c0-132">1.0</span></span>|
|[<span data-ttu-id="eb9c0-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eb9c0-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb9c0-134">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="eb9c0-134">Compose or Read</span></span>|

#### <a name="coerciontype-string"></a><span data-ttu-id="eb9c0-135">CoercionType: chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-135">CoercionType: String</span></span>

<span data-ttu-id="eb9c0-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eb9c0-137">Type</span><span class="sxs-lookup"><span data-stu-id="eb9c0-137">Type</span></span>

*   <span data-ttu-id="eb9c0-138">String</span><span class="sxs-lookup"><span data-stu-id="eb9c0-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eb9c0-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="eb9c0-139">Properties:</span></span>

|<span data-ttu-id="eb9c0-140">Nom</span><span class="sxs-lookup"><span data-stu-id="eb9c0-140">Name</span></span>| <span data-ttu-id="eb9c0-141">Type</span><span class="sxs-lookup"><span data-stu-id="eb9c0-141">Type</span></span>| <span data-ttu-id="eb9c0-142">Description</span><span class="sxs-lookup"><span data-stu-id="eb9c0-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="eb9c0-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-143">String</span></span>|<span data-ttu-id="eb9c0-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="eb9c0-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-145">String</span></span>|<span data-ttu-id="eb9c0-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb9c0-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eb9c0-147">Requirements</span></span>

|<span data-ttu-id="eb9c0-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eb9c0-148">Requirement</span></span>| <span data-ttu-id="eb9c0-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="eb9c0-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb9c0-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eb9c0-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb9c0-151">1.0</span><span class="sxs-lookup"><span data-stu-id="eb9c0-151">1.0</span></span>|
|[<span data-ttu-id="eb9c0-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eb9c0-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb9c0-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="eb9c0-153">Compose or Read</span></span>|

#### <a name="sourceproperty-string"></a><span data-ttu-id="eb9c0-154">SourceProperty: chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-154">SourceProperty: String</span></span>

<span data-ttu-id="eb9c0-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="eb9c0-156">Type</span><span class="sxs-lookup"><span data-stu-id="eb9c0-156">Type</span></span>

*   <span data-ttu-id="eb9c0-157">String</span><span class="sxs-lookup"><span data-stu-id="eb9c0-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="eb9c0-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="eb9c0-158">Properties:</span></span>

|<span data-ttu-id="eb9c0-159">Nom</span><span class="sxs-lookup"><span data-stu-id="eb9c0-159">Name</span></span>| <span data-ttu-id="eb9c0-160">Type</span><span class="sxs-lookup"><span data-stu-id="eb9c0-160">Type</span></span>| <span data-ttu-id="eb9c0-161">Description</span><span class="sxs-lookup"><span data-stu-id="eb9c0-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="eb9c0-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-162">String</span></span>|<span data-ttu-id="eb9c0-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="eb9c0-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="eb9c0-164">String</span></span>|<span data-ttu-id="eb9c0-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="eb9c0-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eb9c0-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="eb9c0-166">Requirements</span></span>

|<span data-ttu-id="eb9c0-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="eb9c0-167">Requirement</span></span>| <span data-ttu-id="eb9c0-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="eb9c0-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb9c0-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="eb9c0-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eb9c0-170">1.0</span><span class="sxs-lookup"><span data-stu-id="eb9c0-170">1.0</span></span>|
|[<span data-ttu-id="eb9c0-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="eb9c0-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eb9c0-172">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="eb9c0-172">Compose or Read</span></span>|
