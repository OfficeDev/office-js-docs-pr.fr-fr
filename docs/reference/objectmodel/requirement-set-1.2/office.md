---
title: Espace de noms Office-ensemble de conditions requises 1,2
description: Modèle objet pour l’espace de noms de niveau supérieur de l’API des compléments Outlook (version 1,2 de l’API de boîte aux lettres).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 10445204d3007d816ebed74ede9eeab5d3dfd83c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720161"
---
# <a name="office"></a><span data-ttu-id="bc230-103">Office</span><span class="sxs-lookup"><span data-stu-id="bc230-103">Office</span></span>

<span data-ttu-id="bc230-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="bc230-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bc230-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc230-106">Requirements</span></span>

|<span data-ttu-id="bc230-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc230-107">Requirement</span></span>| <span data-ttu-id="bc230-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc230-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc230-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc230-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bc230-110">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-110">1.1</span></span>|
|[<span data-ttu-id="bc230-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc230-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bc230-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="bc230-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="bc230-113">Properties</span></span>

| <span data-ttu-id="bc230-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="bc230-114">Property</span></span> | <span data-ttu-id="bc230-115">Modes</span><span class="sxs-lookup"><span data-stu-id="bc230-115">Modes</span></span> | <span data-ttu-id="bc230-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="bc230-116">Return type</span></span> | <span data-ttu-id="bc230-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="bc230-117">Minimum</span></span><br><span data-ttu-id="bc230-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc230-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bc230-119">context</span><span class="sxs-lookup"><span data-stu-id="bc230-119">context</span></span>](office.context.md) | <span data-ttu-id="bc230-120">Composition</span><span class="sxs-lookup"><span data-stu-id="bc230-120">Compose</span></span><br><span data-ttu-id="bc230-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-121">Read</span></span> | [<span data-ttu-id="bc230-122">Context</span><span class="sxs-lookup"><span data-stu-id="bc230-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2) | [<span data-ttu-id="bc230-123">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="bc230-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="bc230-124">Enumerations</span></span>

| <span data-ttu-id="bc230-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="bc230-125">Enumeration</span></span> | <span data-ttu-id="bc230-126">Modes</span><span class="sxs-lookup"><span data-stu-id="bc230-126">Modes</span></span> | <span data-ttu-id="bc230-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="bc230-127">Return type</span></span> | <span data-ttu-id="bc230-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="bc230-128">Minimum</span></span><br><span data-ttu-id="bc230-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc230-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="bc230-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="bc230-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="bc230-131">Composition</span><span class="sxs-lookup"><span data-stu-id="bc230-131">Compose</span></span><br><span data-ttu-id="bc230-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-132">Read</span></span> | <span data-ttu-id="bc230-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-133">String</span></span> | [<span data-ttu-id="bc230-134">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bc230-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="bc230-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="bc230-136">Composition</span><span class="sxs-lookup"><span data-stu-id="bc230-136">Compose</span></span><br><span data-ttu-id="bc230-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-137">Read</span></span> | <span data-ttu-id="bc230-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-138">String</span></span> | [<span data-ttu-id="bc230-139">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="bc230-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="bc230-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="bc230-141">Composition</span><span class="sxs-lookup"><span data-stu-id="bc230-141">Compose</span></span><br><span data-ttu-id="bc230-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-142">Read</span></span> | <span data-ttu-id="bc230-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-143">String</span></span> | [<span data-ttu-id="bc230-144">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="bc230-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="bc230-145">Namespaces</span></span>

<span data-ttu-id="bc230-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="bc230-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="bc230-147">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="bc230-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="bc230-148">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="bc230-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="bc230-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="bc230-150">Type</span><span class="sxs-lookup"><span data-stu-id="bc230-150">Type</span></span>

*   <span data-ttu-id="bc230-151">String</span><span class="sxs-lookup"><span data-stu-id="bc230-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bc230-152">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="bc230-152">Properties:</span></span>

|<span data-ttu-id="bc230-153">Nom</span><span class="sxs-lookup"><span data-stu-id="bc230-153">Name</span></span>| <span data-ttu-id="bc230-154">Type</span><span class="sxs-lookup"><span data-stu-id="bc230-154">Type</span></span>| <span data-ttu-id="bc230-155">Description</span><span class="sxs-lookup"><span data-stu-id="bc230-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="bc230-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-156">String</span></span>|<span data-ttu-id="bc230-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="bc230-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="bc230-158">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-158">String</span></span>|<span data-ttu-id="bc230-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="bc230-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc230-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc230-160">Requirements</span></span>

|<span data-ttu-id="bc230-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc230-161">Requirement</span></span>| <span data-ttu-id="bc230-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc230-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc230-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc230-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bc230-164">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-164">1.1</span></span>|
|[<span data-ttu-id="bc230-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc230-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bc230-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="bc230-167">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-167">CoercionType: String</span></span>

<span data-ttu-id="bc230-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="bc230-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bc230-169">Type</span><span class="sxs-lookup"><span data-stu-id="bc230-169">Type</span></span>

*   <span data-ttu-id="bc230-170">String</span><span class="sxs-lookup"><span data-stu-id="bc230-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bc230-171">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="bc230-171">Properties:</span></span>

|<span data-ttu-id="bc230-172">Nom</span><span class="sxs-lookup"><span data-stu-id="bc230-172">Name</span></span>| <span data-ttu-id="bc230-173">Type</span><span class="sxs-lookup"><span data-stu-id="bc230-173">Type</span></span>| <span data-ttu-id="bc230-174">Description</span><span class="sxs-lookup"><span data-stu-id="bc230-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="bc230-175">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-175">String</span></span>|<span data-ttu-id="bc230-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="bc230-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="bc230-177">String</span><span class="sxs-lookup"><span data-stu-id="bc230-177">String</span></span>|<span data-ttu-id="bc230-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="bc230-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc230-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc230-179">Requirements</span></span>

|<span data-ttu-id="bc230-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc230-180">Requirement</span></span>| <span data-ttu-id="bc230-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc230-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc230-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc230-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bc230-183">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-183">1.1</span></span>|
|[<span data-ttu-id="bc230-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc230-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bc230-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="bc230-186">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-186">SourceProperty: String</span></span>

<span data-ttu-id="bc230-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="bc230-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="bc230-188">Type</span><span class="sxs-lookup"><span data-stu-id="bc230-188">Type</span></span>

*   <span data-ttu-id="bc230-189">String</span><span class="sxs-lookup"><span data-stu-id="bc230-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="bc230-190">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="bc230-190">Properties:</span></span>

|<span data-ttu-id="bc230-191">Nom</span><span class="sxs-lookup"><span data-stu-id="bc230-191">Name</span></span>| <span data-ttu-id="bc230-192">Type</span><span class="sxs-lookup"><span data-stu-id="bc230-192">Type</span></span>| <span data-ttu-id="bc230-193">Description</span><span class="sxs-lookup"><span data-stu-id="bc230-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="bc230-194">Chaîne</span><span class="sxs-lookup"><span data-stu-id="bc230-194">String</span></span>|<span data-ttu-id="bc230-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="bc230-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="bc230-196">String</span><span class="sxs-lookup"><span data-stu-id="bc230-196">String</span></span>|<span data-ttu-id="bc230-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="bc230-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bc230-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="bc230-198">Requirements</span></span>

|<span data-ttu-id="bc230-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="bc230-199">Requirement</span></span>| <span data-ttu-id="bc230-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="bc230-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="bc230-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="bc230-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="bc230-202">1.1</span><span class="sxs-lookup"><span data-stu-id="bc230-202">1.1</span></span>|
|[<span data-ttu-id="bc230-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="bc230-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="bc230-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="bc230-204">Compose or Read</span></span>|
