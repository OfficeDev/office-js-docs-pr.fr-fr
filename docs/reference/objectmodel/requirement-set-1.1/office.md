---
title: Espace de noms Office-ensemble de conditions requises 1,1
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,1.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: f127ab6594d5838700bbc04661d995b01da4f067
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611476"
---
# <a name="office-mailbox-requirement-set-11"></a><span data-ttu-id="7deaf-103">Office (boîte aux lettres requise définie sur 1,1)</span><span class="sxs-lookup"><span data-stu-id="7deaf-103">Office (Mailbox requirement set 1.1)</span></span>

<span data-ttu-id="7deaf-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="7deaf-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7deaf-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7deaf-106">Requirements</span></span>

|<span data-ttu-id="7deaf-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7deaf-107">Requirement</span></span>| <span data-ttu-id="7deaf-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="7deaf-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="7deaf-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7deaf-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7deaf-110">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-110">1.1</span></span>|
|[<span data-ttu-id="7deaf-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7deaf-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7deaf-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7deaf-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="7deaf-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="7deaf-113">Properties</span></span>

| <span data-ttu-id="7deaf-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="7deaf-114">Property</span></span> | <span data-ttu-id="7deaf-115">Modes</span><span class="sxs-lookup"><span data-stu-id="7deaf-115">Modes</span></span> | <span data-ttu-id="7deaf-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="7deaf-116">Return type</span></span> | <span data-ttu-id="7deaf-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="7deaf-117">Minimum</span></span><br><span data-ttu-id="7deaf-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="7deaf-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7deaf-119">context</span><span class="sxs-lookup"><span data-stu-id="7deaf-119">context</span></span>](office.context.md) | <span data-ttu-id="7deaf-120">Composition</span><span class="sxs-lookup"><span data-stu-id="7deaf-120">Compose</span></span><br><span data-ttu-id="7deaf-121">Read</span><span class="sxs-lookup"><span data-stu-id="7deaf-121">Read</span></span> | [<span data-ttu-id="7deaf-122">Context</span><span class="sxs-lookup"><span data-stu-id="7deaf-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.1) | [<span data-ttu-id="7deaf-123">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="7deaf-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="7deaf-124">Enumerations</span></span>

| <span data-ttu-id="7deaf-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="7deaf-125">Enumeration</span></span> | <span data-ttu-id="7deaf-126">Modes</span><span class="sxs-lookup"><span data-stu-id="7deaf-126">Modes</span></span> | <span data-ttu-id="7deaf-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="7deaf-127">Return type</span></span> | <span data-ttu-id="7deaf-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="7deaf-128">Minimum</span></span><br><span data-ttu-id="7deaf-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="7deaf-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7deaf-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="7deaf-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="7deaf-131">Composition</span><span class="sxs-lookup"><span data-stu-id="7deaf-131">Compose</span></span><br><span data-ttu-id="7deaf-132">Read</span><span class="sxs-lookup"><span data-stu-id="7deaf-132">Read</span></span> | <span data-ttu-id="7deaf-133">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-133">String</span></span> | [<span data-ttu-id="7deaf-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7deaf-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="7deaf-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="7deaf-136">Composition</span><span class="sxs-lookup"><span data-stu-id="7deaf-136">Compose</span></span><br><span data-ttu-id="7deaf-137">Read</span><span class="sxs-lookup"><span data-stu-id="7deaf-137">Read</span></span> | <span data-ttu-id="7deaf-138">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-138">String</span></span> | [<span data-ttu-id="7deaf-139">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7deaf-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="7deaf-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="7deaf-141">Composition</span><span class="sxs-lookup"><span data-stu-id="7deaf-141">Compose</span></span><br><span data-ttu-id="7deaf-142">Read</span><span class="sxs-lookup"><span data-stu-id="7deaf-142">Read</span></span> | <span data-ttu-id="7deaf-143">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-143">String</span></span> | [<span data-ttu-id="7deaf-144">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="7deaf-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="7deaf-145">Namespaces</span></span>

<span data-ttu-id="7deaf-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="7deaf-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.1): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="7deaf-147">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="7deaf-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="7deaf-148">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="7deaf-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="7deaf-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="7deaf-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="7deaf-150">Type</span><span class="sxs-lookup"><span data-stu-id="7deaf-150">Type</span></span>

*   <span data-ttu-id="7deaf-151">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7deaf-152">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="7deaf-152">Properties:</span></span>

|<span data-ttu-id="7deaf-153">Nom</span><span class="sxs-lookup"><span data-stu-id="7deaf-153">Name</span></span>| <span data-ttu-id="7deaf-154">Type</span><span class="sxs-lookup"><span data-stu-id="7deaf-154">Type</span></span>| <span data-ttu-id="7deaf-155">Description</span><span class="sxs-lookup"><span data-stu-id="7deaf-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="7deaf-156">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-156">String</span></span>|<span data-ttu-id="7deaf-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="7deaf-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="7deaf-158">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-158">String</span></span>|<span data-ttu-id="7deaf-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="7deaf-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7deaf-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7deaf-160">Requirements</span></span>

|<span data-ttu-id="7deaf-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7deaf-161">Requirement</span></span>| <span data-ttu-id="7deaf-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="7deaf-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="7deaf-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7deaf-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7deaf-164">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-164">1.1</span></span>|
|[<span data-ttu-id="7deaf-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7deaf-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7deaf-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7deaf-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="7deaf-167">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="7deaf-167">CoercionType: String</span></span>

<span data-ttu-id="7deaf-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="7deaf-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7deaf-169">Type</span><span class="sxs-lookup"><span data-stu-id="7deaf-169">Type</span></span>

*   <span data-ttu-id="7deaf-170">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7deaf-171">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="7deaf-171">Properties:</span></span>

|<span data-ttu-id="7deaf-172">Nom</span><span class="sxs-lookup"><span data-stu-id="7deaf-172">Name</span></span>| <span data-ttu-id="7deaf-173">Type</span><span class="sxs-lookup"><span data-stu-id="7deaf-173">Type</span></span>| <span data-ttu-id="7deaf-174">Description</span><span class="sxs-lookup"><span data-stu-id="7deaf-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="7deaf-175">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-175">String</span></span>|<span data-ttu-id="7deaf-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="7deaf-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="7deaf-177">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-177">String</span></span>|<span data-ttu-id="7deaf-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="7deaf-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7deaf-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7deaf-179">Requirements</span></span>

|<span data-ttu-id="7deaf-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7deaf-180">Requirement</span></span>| <span data-ttu-id="7deaf-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="7deaf-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="7deaf-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7deaf-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7deaf-183">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-183">1.1</span></span>|
|[<span data-ttu-id="7deaf-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7deaf-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7deaf-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7deaf-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="7deaf-186">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="7deaf-186">SourceProperty: String</span></span>

<span data-ttu-id="7deaf-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="7deaf-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="7deaf-188">Type</span><span class="sxs-lookup"><span data-stu-id="7deaf-188">Type</span></span>

*   <span data-ttu-id="7deaf-189">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="7deaf-190">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="7deaf-190">Properties:</span></span>

|<span data-ttu-id="7deaf-191">Nom</span><span class="sxs-lookup"><span data-stu-id="7deaf-191">Name</span></span>| <span data-ttu-id="7deaf-192">Type</span><span class="sxs-lookup"><span data-stu-id="7deaf-192">Type</span></span>| <span data-ttu-id="7deaf-193">Description</span><span class="sxs-lookup"><span data-stu-id="7deaf-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="7deaf-194">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-194">String</span></span>|<span data-ttu-id="7deaf-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="7deaf-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="7deaf-196">String</span><span class="sxs-lookup"><span data-stu-id="7deaf-196">String</span></span>|<span data-ttu-id="7deaf-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="7deaf-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7deaf-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7deaf-198">Requirements</span></span>

|<span data-ttu-id="7deaf-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7deaf-199">Requirement</span></span>| <span data-ttu-id="7deaf-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="7deaf-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="7deaf-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7deaf-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7deaf-202">1.1</span><span class="sxs-lookup"><span data-stu-id="7deaf-202">1.1</span></span>|
|[<span data-ttu-id="7deaf-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7deaf-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7deaf-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7deaf-204">Compose or Read</span></span>|