---
title: Espace de noms Office-ensemble de conditions requises 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: eda5e1fb5f2c11ae91e4a1479892c36ec23e1897
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871996"
---
# <a name="office"></a><span data-ttu-id="d7980-102">Office</span><span class="sxs-lookup"><span data-stu-id="d7980-102">Office</span></span>

<span data-ttu-id="d7980-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d7980-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7980-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7980-105">Requirements</span></span>

|<span data-ttu-id="d7980-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7980-106">Requirement</span></span>| <span data-ttu-id="d7980-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7980-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7980-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7980-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7980-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d7980-109">1.0</span></span>|
|[<span data-ttu-id="d7980-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7980-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7980-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7980-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="d7980-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="d7980-112">Namespaces</span></span>

<span data-ttu-id="d7980-113">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="d7980-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d7980-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="d7980-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d7980-115">Membres</span><span class="sxs-lookup"><span data-stu-id="d7980-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d7980-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d7980-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="d7980-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d7980-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d7980-118">Type</span><span class="sxs-lookup"><span data-stu-id="d7980-118">Type</span></span>

*   <span data-ttu-id="d7980-119">String</span><span class="sxs-lookup"><span data-stu-id="d7980-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7980-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d7980-120">Properties:</span></span>

|<span data-ttu-id="d7980-121">Nom</span><span class="sxs-lookup"><span data-stu-id="d7980-121">Name</span></span>| <span data-ttu-id="d7980-122">Type</span><span class="sxs-lookup"><span data-stu-id="d7980-122">Type</span></span>| <span data-ttu-id="d7980-123">Description</span><span class="sxs-lookup"><span data-stu-id="d7980-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d7980-124">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7980-124">String</span></span>|<span data-ttu-id="d7980-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="d7980-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d7980-126">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7980-126">String</span></span>|<span data-ttu-id="d7980-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="d7980-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7980-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7980-128">Requirements</span></span>

|<span data-ttu-id="d7980-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7980-129">Requirement</span></span>| <span data-ttu-id="d7980-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7980-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7980-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7980-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7980-132">1.0</span><span class="sxs-lookup"><span data-stu-id="d7980-132">1.0</span></span>|
|[<span data-ttu-id="d7980-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7980-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7980-134">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7980-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="d7980-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d7980-135">CoercionType :String</span></span>

<span data-ttu-id="d7980-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d7980-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7980-137">Type</span><span class="sxs-lookup"><span data-stu-id="d7980-137">Type</span></span>

*   <span data-ttu-id="d7980-138">String</span><span class="sxs-lookup"><span data-stu-id="d7980-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7980-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d7980-139">Properties:</span></span>

|<span data-ttu-id="d7980-140">Nom</span><span class="sxs-lookup"><span data-stu-id="d7980-140">Name</span></span>| <span data-ttu-id="d7980-141">Type</span><span class="sxs-lookup"><span data-stu-id="d7980-141">Type</span></span>| <span data-ttu-id="d7980-142">Description</span><span class="sxs-lookup"><span data-stu-id="d7980-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d7980-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7980-143">String</span></span>|<span data-ttu-id="d7980-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="d7980-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d7980-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7980-145">String</span></span>|<span data-ttu-id="d7980-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="d7980-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7980-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7980-147">Requirements</span></span>

|<span data-ttu-id="d7980-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7980-148">Requirement</span></span>| <span data-ttu-id="d7980-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7980-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7980-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7980-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7980-151">1.0</span><span class="sxs-lookup"><span data-stu-id="d7980-151">1.0</span></span>|
|[<span data-ttu-id="d7980-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7980-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7980-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7980-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="d7980-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d7980-154">SourceProperty :String</span></span>

<span data-ttu-id="d7980-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d7980-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7980-156">Type</span><span class="sxs-lookup"><span data-stu-id="d7980-156">Type</span></span>

*   <span data-ttu-id="d7980-157">String</span><span class="sxs-lookup"><span data-stu-id="d7980-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7980-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d7980-158">Properties:</span></span>

|<span data-ttu-id="d7980-159">Nom</span><span class="sxs-lookup"><span data-stu-id="d7980-159">Name</span></span>| <span data-ttu-id="d7980-160">Type</span><span class="sxs-lookup"><span data-stu-id="d7980-160">Type</span></span>| <span data-ttu-id="d7980-161">Description</span><span class="sxs-lookup"><span data-stu-id="d7980-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d7980-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7980-162">String</span></span>|<span data-ttu-id="d7980-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="d7980-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d7980-164">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7980-164">String</span></span>|<span data-ttu-id="d7980-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="d7980-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7980-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7980-166">Requirements</span></span>

|<span data-ttu-id="d7980-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7980-167">Requirement</span></span>| <span data-ttu-id="d7980-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7980-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7980-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7980-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7980-170">1.0</span><span class="sxs-lookup"><span data-stu-id="d7980-170">1.0</span></span>|
|[<span data-ttu-id="d7980-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7980-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d7980-172">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7980-172">Compose or Read</span></span>|
