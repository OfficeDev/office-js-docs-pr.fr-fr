---
title: Espace de noms Office-ensemble de conditions requises 1,3
description: Modèle objet pour l’espace de noms de niveau supérieur de l’API des compléments Outlook (version 1,3 de l’API de boîte aux lettres).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 706f12f4425a883f0d18fcd6f9ee18972972d72b
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717774"
---
# <a name="office"></a><span data-ttu-id="c553c-103">Office</span><span class="sxs-lookup"><span data-stu-id="c553c-103">Office</span></span>

<span data-ttu-id="c553c-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c553c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c553c-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c553c-106">Requirements</span></span>

|<span data-ttu-id="c553c-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c553c-107">Requirement</span></span>| <span data-ttu-id="c553c-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="c553c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c553c-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c553c-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c553c-110">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-110">1.1</span></span>|
|[<span data-ttu-id="c553c-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c553c-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c553c-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c553c-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="c553c-113">Properties</span></span>

| <span data-ttu-id="c553c-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="c553c-114">Property</span></span> | <span data-ttu-id="c553c-115">Modes</span><span class="sxs-lookup"><span data-stu-id="c553c-115">Modes</span></span> | <span data-ttu-id="c553c-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="c553c-116">Return type</span></span> | <span data-ttu-id="c553c-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="c553c-117">Minimum</span></span><br><span data-ttu-id="c553c-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="c553c-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c553c-119">context</span><span class="sxs-lookup"><span data-stu-id="c553c-119">context</span></span>](office.context.md) | <span data-ttu-id="c553c-120">Composition</span><span class="sxs-lookup"><span data-stu-id="c553c-120">Compose</span></span><br><span data-ttu-id="c553c-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-121">Read</span></span> | [<span data-ttu-id="c553c-122">Context</span><span class="sxs-lookup"><span data-stu-id="c553c-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3) | [<span data-ttu-id="c553c-123">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="c553c-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="c553c-124">Enumerations</span></span>

| <span data-ttu-id="c553c-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="c553c-125">Enumeration</span></span> | <span data-ttu-id="c553c-126">Modes</span><span class="sxs-lookup"><span data-stu-id="c553c-126">Modes</span></span> | <span data-ttu-id="c553c-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="c553c-127">Return type</span></span> | <span data-ttu-id="c553c-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="c553c-128">Minimum</span></span><br><span data-ttu-id="c553c-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="c553c-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c553c-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c553c-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c553c-131">Composition</span><span class="sxs-lookup"><span data-stu-id="c553c-131">Compose</span></span><br><span data-ttu-id="c553c-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-132">Read</span></span> | <span data-ttu-id="c553c-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-133">String</span></span> | [<span data-ttu-id="c553c-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c553c-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c553c-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c553c-136">Composition</span><span class="sxs-lookup"><span data-stu-id="c553c-136">Compose</span></span><br><span data-ttu-id="c553c-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-137">Read</span></span> | <span data-ttu-id="c553c-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-138">String</span></span> | [<span data-ttu-id="c553c-139">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c553c-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c553c-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c553c-141">Composition</span><span class="sxs-lookup"><span data-stu-id="c553c-141">Compose</span></span><br><span data-ttu-id="c553c-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-142">Read</span></span> | <span data-ttu-id="c553c-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-143">String</span></span> | [<span data-ttu-id="c553c-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="c553c-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c553c-145">Namespaces</span></span>

<span data-ttu-id="c553c-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="c553c-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="c553c-147">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="c553c-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="c553c-148">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="c553c-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c553c-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c553c-150">Type</span><span class="sxs-lookup"><span data-stu-id="c553c-150">Type</span></span>

*   <span data-ttu-id="c553c-151">String</span><span class="sxs-lookup"><span data-stu-id="c553c-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c553c-152">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c553c-152">Properties:</span></span>

|<span data-ttu-id="c553c-153">Nom</span><span class="sxs-lookup"><span data-stu-id="c553c-153">Name</span></span>| <span data-ttu-id="c553c-154">Type</span><span class="sxs-lookup"><span data-stu-id="c553c-154">Type</span></span>| <span data-ttu-id="c553c-155">Description</span><span class="sxs-lookup"><span data-stu-id="c553c-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c553c-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-156">String</span></span>|<span data-ttu-id="c553c-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="c553c-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c553c-158">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-158">String</span></span>|<span data-ttu-id="c553c-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="c553c-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c553c-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c553c-160">Requirements</span></span>

|<span data-ttu-id="c553c-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c553c-161">Requirement</span></span>| <span data-ttu-id="c553c-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="c553c-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="c553c-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c553c-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c553c-164">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-164">1.1</span></span>|
|[<span data-ttu-id="c553c-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c553c-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c553c-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="c553c-167">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-167">CoercionType: String</span></span>

<span data-ttu-id="c553c-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c553c-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c553c-169">Type</span><span class="sxs-lookup"><span data-stu-id="c553c-169">Type</span></span>

*   <span data-ttu-id="c553c-170">String</span><span class="sxs-lookup"><span data-stu-id="c553c-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c553c-171">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c553c-171">Properties:</span></span>

|<span data-ttu-id="c553c-172">Nom</span><span class="sxs-lookup"><span data-stu-id="c553c-172">Name</span></span>| <span data-ttu-id="c553c-173">Type</span><span class="sxs-lookup"><span data-stu-id="c553c-173">Type</span></span>| <span data-ttu-id="c553c-174">Description</span><span class="sxs-lookup"><span data-stu-id="c553c-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c553c-175">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-175">String</span></span>|<span data-ttu-id="c553c-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="c553c-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c553c-177">String</span><span class="sxs-lookup"><span data-stu-id="c553c-177">String</span></span>|<span data-ttu-id="c553c-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="c553c-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c553c-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c553c-179">Requirements</span></span>

|<span data-ttu-id="c553c-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c553c-180">Requirement</span></span>| <span data-ttu-id="c553c-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="c553c-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c553c-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c553c-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c553c-183">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-183">1.1</span></span>|
|[<span data-ttu-id="c553c-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c553c-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c553c-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="c553c-186">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-186">SourceProperty: String</span></span>

<span data-ttu-id="c553c-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c553c-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c553c-188">Type</span><span class="sxs-lookup"><span data-stu-id="c553c-188">Type</span></span>

*   <span data-ttu-id="c553c-189">String</span><span class="sxs-lookup"><span data-stu-id="c553c-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c553c-190">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c553c-190">Properties:</span></span>

|<span data-ttu-id="c553c-191">Nom</span><span class="sxs-lookup"><span data-stu-id="c553c-191">Name</span></span>| <span data-ttu-id="c553c-192">Type</span><span class="sxs-lookup"><span data-stu-id="c553c-192">Type</span></span>| <span data-ttu-id="c553c-193">Description</span><span class="sxs-lookup"><span data-stu-id="c553c-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c553c-194">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c553c-194">String</span></span>|<span data-ttu-id="c553c-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="c553c-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c553c-196">String</span><span class="sxs-lookup"><span data-stu-id="c553c-196">String</span></span>|<span data-ttu-id="c553c-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="c553c-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c553c-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c553c-198">Requirements</span></span>

|<span data-ttu-id="c553c-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c553c-199">Requirement</span></span>| <span data-ttu-id="c553c-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="c553c-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="c553c-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c553c-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c553c-202">1.1</span><span class="sxs-lookup"><span data-stu-id="c553c-202">1.1</span></span>|
|[<span data-ttu-id="c553c-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c553c-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c553c-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c553c-204">Compose or Read</span></span>|
