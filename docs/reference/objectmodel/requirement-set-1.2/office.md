---
title: Espace de noms Office-ensemble de conditions requises 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: eff7896214866e71b92a1c8a0c72a16e622873f3
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067859"
---
# <a name="office"></a><span data-ttu-id="d13f5-102">Office</span><span class="sxs-lookup"><span data-stu-id="d13f5-102">Office</span></span>

<span data-ttu-id="d13f5-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d13f5-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d13f5-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d13f5-105">Requirements</span></span>

|<span data-ttu-id="d13f5-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d13f5-106">Requirement</span></span>| <span data-ttu-id="d13f5-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="d13f5-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d13f5-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d13f5-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d13f5-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d13f5-109">1.0</span></span>|
|[<span data-ttu-id="d13f5-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d13f5-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d13f5-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d13f5-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="d13f5-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="d13f5-112">Namespaces</span></span>

<span data-ttu-id="d13f5-113">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="d13f5-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d13f5-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="d13f5-114">[MailboxEnums](/javascript/api/outlook_1_2/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d13f5-115">Membres</span><span class="sxs-lookup"><span data-stu-id="d13f5-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d13f5-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d13f5-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="d13f5-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d13f5-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d13f5-118">Type</span><span class="sxs-lookup"><span data-stu-id="d13f5-118">Type</span></span>

*   <span data-ttu-id="d13f5-119">String</span><span class="sxs-lookup"><span data-stu-id="d13f5-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d13f5-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d13f5-120">Properties:</span></span>

|<span data-ttu-id="d13f5-121">Nom</span><span class="sxs-lookup"><span data-stu-id="d13f5-121">Name</span></span>| <span data-ttu-id="d13f5-122">Type</span><span class="sxs-lookup"><span data-stu-id="d13f5-122">Type</span></span>| <span data-ttu-id="d13f5-123">Description</span><span class="sxs-lookup"><span data-stu-id="d13f5-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d13f5-124">String</span><span class="sxs-lookup"><span data-stu-id="d13f5-124">String</span></span>|<span data-ttu-id="d13f5-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="d13f5-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d13f5-126">String</span><span class="sxs-lookup"><span data-stu-id="d13f5-126">String</span></span>|<span data-ttu-id="d13f5-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="d13f5-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d13f5-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d13f5-128">Requirements</span></span>

|<span data-ttu-id="d13f5-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d13f5-129">Requirement</span></span>| <span data-ttu-id="d13f5-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="d13f5-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="d13f5-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d13f5-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d13f5-132">1.0</span><span class="sxs-lookup"><span data-stu-id="d13f5-132">1.0</span></span>|
|[<span data-ttu-id="d13f5-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d13f5-133">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d13f5-134">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d13f5-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="d13f5-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d13f5-135">CoercionType :String</span></span>

<span data-ttu-id="d13f5-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d13f5-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d13f5-137">Type</span><span class="sxs-lookup"><span data-stu-id="d13f5-137">Type</span></span>

*   <span data-ttu-id="d13f5-138">String</span><span class="sxs-lookup"><span data-stu-id="d13f5-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d13f5-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d13f5-139">Properties:</span></span>

|<span data-ttu-id="d13f5-140">Nom</span><span class="sxs-lookup"><span data-stu-id="d13f5-140">Name</span></span>| <span data-ttu-id="d13f5-141">Type</span><span class="sxs-lookup"><span data-stu-id="d13f5-141">Type</span></span>| <span data-ttu-id="d13f5-142">Description</span><span class="sxs-lookup"><span data-stu-id="d13f5-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d13f5-143">String</span><span class="sxs-lookup"><span data-stu-id="d13f5-143">String</span></span>|<span data-ttu-id="d13f5-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="d13f5-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d13f5-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d13f5-145">String</span></span>|<span data-ttu-id="d13f5-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="d13f5-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d13f5-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d13f5-147">Requirements</span></span>

|<span data-ttu-id="d13f5-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d13f5-148">Requirement</span></span>| <span data-ttu-id="d13f5-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="d13f5-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="d13f5-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d13f5-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d13f5-151">1.0</span><span class="sxs-lookup"><span data-stu-id="d13f5-151">1.0</span></span>|
|[<span data-ttu-id="d13f5-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d13f5-152">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d13f5-153">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d13f5-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="d13f5-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d13f5-154">SourceProperty :String</span></span>

<span data-ttu-id="d13f5-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d13f5-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d13f5-156">Type</span><span class="sxs-lookup"><span data-stu-id="d13f5-156">Type</span></span>

*   <span data-ttu-id="d13f5-157">String</span><span class="sxs-lookup"><span data-stu-id="d13f5-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d13f5-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d13f5-158">Properties:</span></span>

|<span data-ttu-id="d13f5-159">Nom</span><span class="sxs-lookup"><span data-stu-id="d13f5-159">Name</span></span>| <span data-ttu-id="d13f5-160">Type</span><span class="sxs-lookup"><span data-stu-id="d13f5-160">Type</span></span>| <span data-ttu-id="d13f5-161">Description</span><span class="sxs-lookup"><span data-stu-id="d13f5-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d13f5-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d13f5-162">String</span></span>|<span data-ttu-id="d13f5-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="d13f5-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d13f5-164">String</span><span class="sxs-lookup"><span data-stu-id="d13f5-164">String</span></span>|<span data-ttu-id="d13f5-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="d13f5-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d13f5-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d13f5-166">Requirements</span></span>

|<span data-ttu-id="d13f5-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d13f5-167">Requirement</span></span>| <span data-ttu-id="d13f5-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="d13f5-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="d13f5-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d13f5-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d13f5-170">1.0</span><span class="sxs-lookup"><span data-stu-id="d13f5-170">1.0</span></span>|
|[<span data-ttu-id="d13f5-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d13f5-171">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d13f5-172">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="d13f5-172">Compose or Read</span></span>|
