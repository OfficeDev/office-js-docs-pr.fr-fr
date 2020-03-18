---
title: Espace de noms Office-ensemble de conditions requises 1,4
description: Modèle objet pour l’espace de noms de niveau supérieur de l’API des compléments Outlook (version 1,4 de l’API de boîte aux lettres).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: e5a5c6de5bb87cb32968d9d9d80c621f0acc238d
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720056"
---
# <a name="office"></a><span data-ttu-id="f7d39-103">Office</span><span class="sxs-lookup"><span data-stu-id="f7d39-103">Office</span></span>

<span data-ttu-id="f7d39-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f7d39-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f7d39-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f7d39-106">Requirements</span></span>

|<span data-ttu-id="f7d39-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f7d39-107">Requirement</span></span>| <span data-ttu-id="f7d39-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="f7d39-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7d39-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f7d39-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7d39-110">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-110">1.1</span></span>|
|[<span data-ttu-id="f7d39-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f7d39-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7d39-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f7d39-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="f7d39-113">Properties</span></span>

| <span data-ttu-id="f7d39-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="f7d39-114">Property</span></span> | <span data-ttu-id="f7d39-115">Modes</span><span class="sxs-lookup"><span data-stu-id="f7d39-115">Modes</span></span> | <span data-ttu-id="f7d39-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f7d39-116">Return type</span></span> | <span data-ttu-id="f7d39-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="f7d39-117">Minimum</span></span><br><span data-ttu-id="f7d39-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f7d39-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f7d39-119">context</span><span class="sxs-lookup"><span data-stu-id="f7d39-119">context</span></span>](office.context.md) | <span data-ttu-id="f7d39-120">Composition</span><span class="sxs-lookup"><span data-stu-id="f7d39-120">Compose</span></span><br><span data-ttu-id="f7d39-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-121">Read</span></span> | [<span data-ttu-id="f7d39-122">Context</span><span class="sxs-lookup"><span data-stu-id="f7d39-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="f7d39-123">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f7d39-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="f7d39-124">Enumerations</span></span>

| <span data-ttu-id="f7d39-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="f7d39-125">Enumeration</span></span> | <span data-ttu-id="f7d39-126">Modes</span><span class="sxs-lookup"><span data-stu-id="f7d39-126">Modes</span></span> | <span data-ttu-id="f7d39-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f7d39-127">Return type</span></span> | <span data-ttu-id="f7d39-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="f7d39-128">Minimum</span></span><br><span data-ttu-id="f7d39-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f7d39-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f7d39-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f7d39-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f7d39-131">Composition</span><span class="sxs-lookup"><span data-stu-id="f7d39-131">Compose</span></span><br><span data-ttu-id="f7d39-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-132">Read</span></span> | <span data-ttu-id="f7d39-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-133">String</span></span> | [<span data-ttu-id="f7d39-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7d39-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f7d39-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f7d39-136">Composition</span><span class="sxs-lookup"><span data-stu-id="f7d39-136">Compose</span></span><br><span data-ttu-id="f7d39-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-137">Read</span></span> | <span data-ttu-id="f7d39-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-138">String</span></span> | [<span data-ttu-id="f7d39-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f7d39-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f7d39-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f7d39-141">Composition</span><span class="sxs-lookup"><span data-stu-id="f7d39-141">Compose</span></span><br><span data-ttu-id="f7d39-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-142">Read</span></span> | <span data-ttu-id="f7d39-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-143">String</span></span> | [<span data-ttu-id="f7d39-144">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f7d39-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="f7d39-145">Namespaces</span></span>

<span data-ttu-id="f7d39-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f7d39-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f7d39-147">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="f7d39-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f7d39-148">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="f7d39-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f7d39-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f7d39-150">Type</span><span class="sxs-lookup"><span data-stu-id="f7d39-150">Type</span></span>

*   <span data-ttu-id="f7d39-151">String</span><span class="sxs-lookup"><span data-stu-id="f7d39-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7d39-152">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f7d39-152">Properties:</span></span>

|<span data-ttu-id="f7d39-153">Nom</span><span class="sxs-lookup"><span data-stu-id="f7d39-153">Name</span></span>| <span data-ttu-id="f7d39-154">Type</span><span class="sxs-lookup"><span data-stu-id="f7d39-154">Type</span></span>| <span data-ttu-id="f7d39-155">Description</span><span class="sxs-lookup"><span data-stu-id="f7d39-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f7d39-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-156">String</span></span>|<span data-ttu-id="f7d39-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="f7d39-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f7d39-158">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-158">String</span></span>|<span data-ttu-id="f7d39-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="f7d39-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7d39-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f7d39-160">Requirements</span></span>

|<span data-ttu-id="f7d39-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f7d39-161">Requirement</span></span>| <span data-ttu-id="f7d39-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="f7d39-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7d39-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f7d39-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7d39-164">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-164">1.1</span></span>|
|[<span data-ttu-id="f7d39-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f7d39-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7d39-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f7d39-167">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-167">CoercionType: String</span></span>

<span data-ttu-id="f7d39-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f7d39-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7d39-169">Type</span><span class="sxs-lookup"><span data-stu-id="f7d39-169">Type</span></span>

*   <span data-ttu-id="f7d39-170">String</span><span class="sxs-lookup"><span data-stu-id="f7d39-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7d39-171">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f7d39-171">Properties:</span></span>

|<span data-ttu-id="f7d39-172">Nom</span><span class="sxs-lookup"><span data-stu-id="f7d39-172">Name</span></span>| <span data-ttu-id="f7d39-173">Type</span><span class="sxs-lookup"><span data-stu-id="f7d39-173">Type</span></span>| <span data-ttu-id="f7d39-174">Description</span><span class="sxs-lookup"><span data-stu-id="f7d39-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f7d39-175">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-175">String</span></span>|<span data-ttu-id="f7d39-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f7d39-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f7d39-177">String</span><span class="sxs-lookup"><span data-stu-id="f7d39-177">String</span></span>|<span data-ttu-id="f7d39-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="f7d39-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7d39-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f7d39-179">Requirements</span></span>

|<span data-ttu-id="f7d39-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f7d39-180">Requirement</span></span>| <span data-ttu-id="f7d39-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="f7d39-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7d39-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f7d39-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7d39-183">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-183">1.1</span></span>|
|[<span data-ttu-id="f7d39-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f7d39-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7d39-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f7d39-186">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-186">SourceProperty: String</span></span>

<span data-ttu-id="f7d39-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f7d39-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f7d39-188">Type</span><span class="sxs-lookup"><span data-stu-id="f7d39-188">Type</span></span>

*   <span data-ttu-id="f7d39-189">String</span><span class="sxs-lookup"><span data-stu-id="f7d39-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f7d39-190">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f7d39-190">Properties:</span></span>

|<span data-ttu-id="f7d39-191">Nom</span><span class="sxs-lookup"><span data-stu-id="f7d39-191">Name</span></span>| <span data-ttu-id="f7d39-192">Type</span><span class="sxs-lookup"><span data-stu-id="f7d39-192">Type</span></span>| <span data-ttu-id="f7d39-193">Description</span><span class="sxs-lookup"><span data-stu-id="f7d39-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f7d39-194">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f7d39-194">String</span></span>|<span data-ttu-id="f7d39-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f7d39-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f7d39-196">String</span><span class="sxs-lookup"><span data-stu-id="f7d39-196">String</span></span>|<span data-ttu-id="f7d39-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="f7d39-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f7d39-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f7d39-198">Requirements</span></span>

|<span data-ttu-id="f7d39-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f7d39-199">Requirement</span></span>| <span data-ttu-id="f7d39-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="f7d39-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="f7d39-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f7d39-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f7d39-202">1.1</span><span class="sxs-lookup"><span data-stu-id="f7d39-202">1.1</span></span>|
|[<span data-ttu-id="f7d39-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f7d39-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f7d39-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f7d39-204">Compose or Read</span></span>|
