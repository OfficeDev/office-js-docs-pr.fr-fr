---
title: Espace de noms Office-ensemble de conditions requises 1.1
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: b98f8fa01e2cfdf17d6105beab67199a2cec4317
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068020"
---
# <a name="office"></a><span data-ttu-id="61d7a-102">Office</span><span class="sxs-lookup"><span data-stu-id="61d7a-102">Office</span></span>

<span data-ttu-id="61d7a-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="61d7a-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="61d7a-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="61d7a-105">Requirements</span></span>

|<span data-ttu-id="61d7a-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="61d7a-106">Requirement</span></span>| <span data-ttu-id="61d7a-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="61d7a-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="61d7a-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="61d7a-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61d7a-109">1.0</span><span class="sxs-lookup"><span data-stu-id="61d7a-109">1.0</span></span>|
|[<span data-ttu-id="61d7a-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="61d7a-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61d7a-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="61d7a-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="61d7a-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="61d7a-112">Namespaces</span></span>

<span data-ttu-id="61d7a-113">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="61d7a-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="61d7a-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="61d7a-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="61d7a-115">Membres</span><span class="sxs-lookup"><span data-stu-id="61d7a-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="61d7a-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="61d7a-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="61d7a-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="61d7a-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="61d7a-118">Type</span><span class="sxs-lookup"><span data-stu-id="61d7a-118">Type</span></span>

*   <span data-ttu-id="61d7a-119">String</span><span class="sxs-lookup"><span data-stu-id="61d7a-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="61d7a-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="61d7a-120">Properties:</span></span>

|<span data-ttu-id="61d7a-121">Nom</span><span class="sxs-lookup"><span data-stu-id="61d7a-121">Name</span></span>| <span data-ttu-id="61d7a-122">Type</span><span class="sxs-lookup"><span data-stu-id="61d7a-122">Type</span></span>| <span data-ttu-id="61d7a-123">Description</span><span class="sxs-lookup"><span data-stu-id="61d7a-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="61d7a-124">String</span><span class="sxs-lookup"><span data-stu-id="61d7a-124">String</span></span>|<span data-ttu-id="61d7a-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="61d7a-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="61d7a-126">String</span><span class="sxs-lookup"><span data-stu-id="61d7a-126">String</span></span>|<span data-ttu-id="61d7a-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="61d7a-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61d7a-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="61d7a-128">Requirements</span></span>

|<span data-ttu-id="61d7a-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="61d7a-129">Requirement</span></span>| <span data-ttu-id="61d7a-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="61d7a-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="61d7a-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="61d7a-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61d7a-132">1.0</span><span class="sxs-lookup"><span data-stu-id="61d7a-132">1.0</span></span>|
|[<span data-ttu-id="61d7a-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="61d7a-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61d7a-134">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="61d7a-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="61d7a-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="61d7a-135">CoercionType :String</span></span>

<span data-ttu-id="61d7a-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="61d7a-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="61d7a-137">Type</span><span class="sxs-lookup"><span data-stu-id="61d7a-137">Type</span></span>

*   <span data-ttu-id="61d7a-138">String</span><span class="sxs-lookup"><span data-stu-id="61d7a-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="61d7a-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="61d7a-139">Properties:</span></span>

|<span data-ttu-id="61d7a-140">Nom</span><span class="sxs-lookup"><span data-stu-id="61d7a-140">Name</span></span>| <span data-ttu-id="61d7a-141">Type</span><span class="sxs-lookup"><span data-stu-id="61d7a-141">Type</span></span>| <span data-ttu-id="61d7a-142">Description</span><span class="sxs-lookup"><span data-stu-id="61d7a-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="61d7a-143">String</span><span class="sxs-lookup"><span data-stu-id="61d7a-143">String</span></span>|<span data-ttu-id="61d7a-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="61d7a-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="61d7a-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="61d7a-145">String</span></span>|<span data-ttu-id="61d7a-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="61d7a-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61d7a-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="61d7a-147">Requirements</span></span>

|<span data-ttu-id="61d7a-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="61d7a-148">Requirement</span></span>| <span data-ttu-id="61d7a-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="61d7a-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="61d7a-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="61d7a-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61d7a-151">1.0</span><span class="sxs-lookup"><span data-stu-id="61d7a-151">1.0</span></span>|
|[<span data-ttu-id="61d7a-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="61d7a-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61d7a-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="61d7a-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="61d7a-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="61d7a-154">SourceProperty :String</span></span>

<span data-ttu-id="61d7a-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="61d7a-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="61d7a-156">Type</span><span class="sxs-lookup"><span data-stu-id="61d7a-156">Type</span></span>

*   <span data-ttu-id="61d7a-157">String</span><span class="sxs-lookup"><span data-stu-id="61d7a-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="61d7a-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="61d7a-158">Properties:</span></span>

|<span data-ttu-id="61d7a-159">Nom</span><span class="sxs-lookup"><span data-stu-id="61d7a-159">Name</span></span>| <span data-ttu-id="61d7a-160">Type</span><span class="sxs-lookup"><span data-stu-id="61d7a-160">Type</span></span>| <span data-ttu-id="61d7a-161">Description</span><span class="sxs-lookup"><span data-stu-id="61d7a-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="61d7a-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="61d7a-162">String</span></span>|<span data-ttu-id="61d7a-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="61d7a-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="61d7a-164">String</span><span class="sxs-lookup"><span data-stu-id="61d7a-164">String</span></span>|<span data-ttu-id="61d7a-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="61d7a-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="61d7a-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="61d7a-166">Requirements</span></span>

|<span data-ttu-id="61d7a-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="61d7a-167">Requirement</span></span>| <span data-ttu-id="61d7a-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="61d7a-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="61d7a-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="61d7a-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="61d7a-170">1.0</span><span class="sxs-lookup"><span data-stu-id="61d7a-170">1.0</span></span>|
|[<span data-ttu-id="61d7a-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="61d7a-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="61d7a-172">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="61d7a-172">Compose or Read</span></span>|
