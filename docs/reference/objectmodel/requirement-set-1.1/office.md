---
title: Espace de noms Office-ensemble de conditions requises 1,1
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: eda5e1fb5f2c11ae91e4a1479892c36ec23e1897
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451912"
---
# <a name="office"></a><span data-ttu-id="f5a31-102">Office</span><span class="sxs-lookup"><span data-stu-id="f5a31-102">Office</span></span>

<span data-ttu-id="f5a31-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f5a31-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f5a31-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f5a31-105">Requirements</span></span>

|<span data-ttu-id="f5a31-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f5a31-106">Requirement</span></span>| <span data-ttu-id="f5a31-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="f5a31-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5a31-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f5a31-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5a31-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f5a31-109">1.0</span></span>|
|[<span data-ttu-id="f5a31-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f5a31-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5a31-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f5a31-111">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="f5a31-112">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="f5a31-112">Namespaces</span></span>

<span data-ttu-id="f5a31-113">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="f5a31-113">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f5a31-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f5a31-114">[MailboxEnums](/javascript/api/outlook_1_1/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f5a31-115">Membres</span><span class="sxs-lookup"><span data-stu-id="f5a31-115">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f5a31-116">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f5a31-116">AsyncResultStatus :String</span></span>

<span data-ttu-id="f5a31-117">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f5a31-117">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f5a31-118">Type</span><span class="sxs-lookup"><span data-stu-id="f5a31-118">Type</span></span>

*   <span data-ttu-id="f5a31-119">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-119">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f5a31-120">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f5a31-120">Properties:</span></span>

|<span data-ttu-id="f5a31-121">Nom</span><span class="sxs-lookup"><span data-stu-id="f5a31-121">Name</span></span>| <span data-ttu-id="f5a31-122">Type</span><span class="sxs-lookup"><span data-stu-id="f5a31-122">Type</span></span>| <span data-ttu-id="f5a31-123">Description</span><span class="sxs-lookup"><span data-stu-id="f5a31-123">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f5a31-124">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-124">String</span></span>|<span data-ttu-id="f5a31-125">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="f5a31-125">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f5a31-126">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-126">String</span></span>|<span data-ttu-id="f5a31-127">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="f5a31-127">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5a31-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f5a31-128">Requirements</span></span>

|<span data-ttu-id="f5a31-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f5a31-129">Requirement</span></span>| <span data-ttu-id="f5a31-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="f5a31-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5a31-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f5a31-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5a31-132">1.0</span><span class="sxs-lookup"><span data-stu-id="f5a31-132">1.0</span></span>|
|[<span data-ttu-id="f5a31-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f5a31-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5a31-134">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f5a31-134">Compose or Read</span></span>|

####  <a name="coerciontype-string"></a><span data-ttu-id="f5a31-135">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f5a31-135">CoercionType :String</span></span>

<span data-ttu-id="f5a31-136">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f5a31-136">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f5a31-137">Type</span><span class="sxs-lookup"><span data-stu-id="f5a31-137">Type</span></span>

*   <span data-ttu-id="f5a31-138">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-138">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f5a31-139">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f5a31-139">Properties:</span></span>

|<span data-ttu-id="f5a31-140">Nom</span><span class="sxs-lookup"><span data-stu-id="f5a31-140">Name</span></span>| <span data-ttu-id="f5a31-141">Type</span><span class="sxs-lookup"><span data-stu-id="f5a31-141">Type</span></span>| <span data-ttu-id="f5a31-142">Description</span><span class="sxs-lookup"><span data-stu-id="f5a31-142">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f5a31-143">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-143">String</span></span>|<span data-ttu-id="f5a31-144">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f5a31-144">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f5a31-145">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f5a31-145">String</span></span>|<span data-ttu-id="f5a31-146">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="f5a31-146">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5a31-147">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f5a31-147">Requirements</span></span>

|<span data-ttu-id="f5a31-148">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f5a31-148">Requirement</span></span>| <span data-ttu-id="f5a31-149">Valeur</span><span class="sxs-lookup"><span data-stu-id="f5a31-149">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5a31-150">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f5a31-150">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5a31-151">1.0</span><span class="sxs-lookup"><span data-stu-id="f5a31-151">1.0</span></span>|
|[<span data-ttu-id="f5a31-152">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f5a31-152">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5a31-153">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f5a31-153">Compose or Read</span></span>|

####  <a name="sourceproperty-string"></a><span data-ttu-id="f5a31-154">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f5a31-154">SourceProperty :String</span></span>

<span data-ttu-id="f5a31-155">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f5a31-155">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f5a31-156">Type</span><span class="sxs-lookup"><span data-stu-id="f5a31-156">Type</span></span>

*   <span data-ttu-id="f5a31-157">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-157">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f5a31-158">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f5a31-158">Properties:</span></span>

|<span data-ttu-id="f5a31-159">Nom</span><span class="sxs-lookup"><span data-stu-id="f5a31-159">Name</span></span>| <span data-ttu-id="f5a31-160">Type</span><span class="sxs-lookup"><span data-stu-id="f5a31-160">Type</span></span>| <span data-ttu-id="f5a31-161">Description</span><span class="sxs-lookup"><span data-stu-id="f5a31-161">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f5a31-162">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-162">String</span></span>|<span data-ttu-id="f5a31-163">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f5a31-163">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f5a31-164">String</span><span class="sxs-lookup"><span data-stu-id="f5a31-164">String</span></span>|<span data-ttu-id="f5a31-165">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="f5a31-165">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5a31-166">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f5a31-166">Requirements</span></span>

|<span data-ttu-id="f5a31-167">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f5a31-167">Requirement</span></span>| <span data-ttu-id="f5a31-168">Valeur</span><span class="sxs-lookup"><span data-stu-id="f5a31-168">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5a31-169">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f5a31-169">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5a31-170">1.0</span><span class="sxs-lookup"><span data-stu-id="f5a31-170">1.0</span></span>|
|[<span data-ttu-id="f5a31-171">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f5a31-171">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5a31-172">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f5a31-172">Compose or Read</span></span>|
