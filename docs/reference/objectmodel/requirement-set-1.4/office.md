---
title: Espace de noms Office-ensemble de conditions requises 1,4
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: eb49c54a3ec4035c862181aa8e143dcb2d948693
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165411"
---
# <a name="office"></a><span data-ttu-id="ecdb1-102">Office</span><span class="sxs-lookup"><span data-stu-id="ecdb1-102">Office</span></span>

<span data-ttu-id="ecdb1-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ecdb1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ecdb1-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ecdb1-105">Requirements</span></span>

|<span data-ttu-id="ecdb1-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ecdb1-106">Requirement</span></span>| <span data-ttu-id="ecdb1-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="ecdb1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ecdb1-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ecdb1-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ecdb1-109">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-109">1.1</span></span>|
|[<span data-ttu-id="ecdb1-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ecdb1-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ecdb1-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ecdb1-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ecdb1-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ecdb1-112">Properties</span></span>

| <span data-ttu-id="ecdb1-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="ecdb1-113">Property</span></span> | <span data-ttu-id="ecdb1-114">Modes</span><span class="sxs-lookup"><span data-stu-id="ecdb1-114">Modes</span></span> | <span data-ttu-id="ecdb1-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ecdb1-115">Return type</span></span> | <span data-ttu-id="ecdb1-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="ecdb1-116">Minimum</span></span><br><span data-ttu-id="ecdb1-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ecdb1-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ecdb1-118">context</span><span class="sxs-lookup"><span data-stu-id="ecdb1-118">context</span></span>](office.context.md) | <span data-ttu-id="ecdb1-119">Composition</span><span class="sxs-lookup"><span data-stu-id="ecdb1-119">Compose</span></span><br><span data-ttu-id="ecdb1-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="ecdb1-120">Read</span></span> | [<span data-ttu-id="ecdb1-121">Context</span><span class="sxs-lookup"><span data-stu-id="ecdb1-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="ecdb1-122">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="ecdb1-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="ecdb1-123">Enumerations</span></span>

| <span data-ttu-id="ecdb1-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="ecdb1-124">Enumeration</span></span> | <span data-ttu-id="ecdb1-125">Modes</span><span class="sxs-lookup"><span data-stu-id="ecdb1-125">Modes</span></span> | <span data-ttu-id="ecdb1-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ecdb1-126">Return type</span></span> | <span data-ttu-id="ecdb1-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="ecdb1-127">Minimum</span></span><br><span data-ttu-id="ecdb1-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ecdb1-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ecdb1-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ecdb1-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ecdb1-130">Composition</span><span class="sxs-lookup"><span data-stu-id="ecdb1-130">Compose</span></span><br><span data-ttu-id="ecdb1-131">Lire</span><span class="sxs-lookup"><span data-stu-id="ecdb1-131">Read</span></span> | <span data-ttu-id="ecdb1-132">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-132">String</span></span> | [<span data-ttu-id="ecdb1-133">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ecdb1-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ecdb1-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ecdb1-135">Composition</span><span class="sxs-lookup"><span data-stu-id="ecdb1-135">Compose</span></span><br><span data-ttu-id="ecdb1-136">Lire</span><span class="sxs-lookup"><span data-stu-id="ecdb1-136">Read</span></span> | <span data-ttu-id="ecdb1-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-137">String</span></span> | [<span data-ttu-id="ecdb1-138">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ecdb1-139">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ecdb1-139">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ecdb1-140">Composition</span><span class="sxs-lookup"><span data-stu-id="ecdb1-140">Compose</span></span><br><span data-ttu-id="ecdb1-141">Lire</span><span class="sxs-lookup"><span data-stu-id="ecdb1-141">Read</span></span> | <span data-ttu-id="ecdb1-142">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-142">String</span></span> | [<span data-ttu-id="ecdb1-143">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="ecdb1-144">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="ecdb1-144">Namespaces</span></span>

<span data-ttu-id="ecdb1-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-145">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="ecdb1-146">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="ecdb1-146">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="ecdb1-147">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-147">AsyncResultStatus: String</span></span>

<span data-ttu-id="ecdb1-148">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-148">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ecdb1-149">Type</span><span class="sxs-lookup"><span data-stu-id="ecdb1-149">Type</span></span>

*   <span data-ttu-id="ecdb1-150">String</span><span class="sxs-lookup"><span data-stu-id="ecdb1-150">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ecdb1-151">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ecdb1-151">Properties:</span></span>

|<span data-ttu-id="ecdb1-152">Nom</span><span class="sxs-lookup"><span data-stu-id="ecdb1-152">Name</span></span>| <span data-ttu-id="ecdb1-153">Type</span><span class="sxs-lookup"><span data-stu-id="ecdb1-153">Type</span></span>| <span data-ttu-id="ecdb1-154">Description</span><span class="sxs-lookup"><span data-stu-id="ecdb1-154">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ecdb1-155">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-155">String</span></span>|<span data-ttu-id="ecdb1-156">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-156">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ecdb1-157">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-157">String</span></span>|<span data-ttu-id="ecdb1-158">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-158">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ecdb1-159">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ecdb1-159">Requirements</span></span>

|<span data-ttu-id="ecdb1-160">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ecdb1-160">Requirement</span></span>| <span data-ttu-id="ecdb1-161">Valeur</span><span class="sxs-lookup"><span data-stu-id="ecdb1-161">Value</span></span>|
|---|---|
|[<span data-ttu-id="ecdb1-162">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ecdb1-162">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ecdb1-163">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-163">1.1</span></span>|
|[<span data-ttu-id="ecdb1-164">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ecdb1-164">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ecdb1-165">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ecdb1-165">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="ecdb1-166">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-166">CoercionType: String</span></span>

<span data-ttu-id="ecdb1-167">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-167">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ecdb1-168">Type</span><span class="sxs-lookup"><span data-stu-id="ecdb1-168">Type</span></span>

*   <span data-ttu-id="ecdb1-169">String</span><span class="sxs-lookup"><span data-stu-id="ecdb1-169">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ecdb1-170">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ecdb1-170">Properties:</span></span>

|<span data-ttu-id="ecdb1-171">Nom</span><span class="sxs-lookup"><span data-stu-id="ecdb1-171">Name</span></span>| <span data-ttu-id="ecdb1-172">Type</span><span class="sxs-lookup"><span data-stu-id="ecdb1-172">Type</span></span>| <span data-ttu-id="ecdb1-173">Description</span><span class="sxs-lookup"><span data-stu-id="ecdb1-173">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ecdb1-174">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-174">String</span></span>|<span data-ttu-id="ecdb1-175">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-175">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ecdb1-176">String</span><span class="sxs-lookup"><span data-stu-id="ecdb1-176">String</span></span>|<span data-ttu-id="ecdb1-177">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-177">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ecdb1-178">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ecdb1-178">Requirements</span></span>

|<span data-ttu-id="ecdb1-179">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ecdb1-179">Requirement</span></span>| <span data-ttu-id="ecdb1-180">Valeur</span><span class="sxs-lookup"><span data-stu-id="ecdb1-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="ecdb1-181">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ecdb1-181">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ecdb1-182">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-182">1.1</span></span>|
|[<span data-ttu-id="ecdb1-183">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ecdb1-183">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ecdb1-184">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ecdb1-184">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="ecdb1-185">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-185">SourceProperty: String</span></span>

<span data-ttu-id="ecdb1-186">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-186">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ecdb1-187">Type</span><span class="sxs-lookup"><span data-stu-id="ecdb1-187">Type</span></span>

*   <span data-ttu-id="ecdb1-188">String</span><span class="sxs-lookup"><span data-stu-id="ecdb1-188">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ecdb1-189">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ecdb1-189">Properties:</span></span>

|<span data-ttu-id="ecdb1-190">Nom</span><span class="sxs-lookup"><span data-stu-id="ecdb1-190">Name</span></span>| <span data-ttu-id="ecdb1-191">Type</span><span class="sxs-lookup"><span data-stu-id="ecdb1-191">Type</span></span>| <span data-ttu-id="ecdb1-192">Description</span><span class="sxs-lookup"><span data-stu-id="ecdb1-192">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ecdb1-193">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ecdb1-193">String</span></span>|<span data-ttu-id="ecdb1-194">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-194">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ecdb1-195">String</span><span class="sxs-lookup"><span data-stu-id="ecdb1-195">String</span></span>|<span data-ttu-id="ecdb1-196">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="ecdb1-196">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ecdb1-197">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ecdb1-197">Requirements</span></span>

|<span data-ttu-id="ecdb1-198">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ecdb1-198">Requirement</span></span>| <span data-ttu-id="ecdb1-199">Valeur</span><span class="sxs-lookup"><span data-stu-id="ecdb1-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="ecdb1-200">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ecdb1-200">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ecdb1-201">1.1</span><span class="sxs-lookup"><span data-stu-id="ecdb1-201">1.1</span></span>|
|[<span data-ttu-id="ecdb1-202">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ecdb1-202">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ecdb1-203">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ecdb1-203">Compose or Read</span></span>|
