---
title: Espace de noms Office-ensemble de conditions requises 1,2
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,2.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 0dfde315cb71642a995b4c07a1966d3dee3c0d50
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431303"
---
# <a name="office-mailbox-requirement-set-12"></a><span data-ttu-id="51408-103">Office (boîte aux lettres requise définie sur 1,2)</span><span class="sxs-lookup"><span data-stu-id="51408-103">Office (Mailbox requirement set 1.2)</span></span>

<span data-ttu-id="51408-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="51408-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="51408-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="51408-106">Requirements</span></span>

|<span data-ttu-id="51408-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="51408-107">Requirement</span></span>| <span data-ttu-id="51408-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="51408-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="51408-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="51408-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="51408-110">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-110">1.1</span></span>|
|[<span data-ttu-id="51408-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="51408-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="51408-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="51408-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="51408-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="51408-113">Properties</span></span>

| <span data-ttu-id="51408-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="51408-114">Property</span></span> | <span data-ttu-id="51408-115">Modes</span><span class="sxs-lookup"><span data-stu-id="51408-115">Modes</span></span> | <span data-ttu-id="51408-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="51408-116">Return type</span></span> | <span data-ttu-id="51408-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="51408-117">Minimum</span></span><br><span data-ttu-id="51408-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="51408-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="51408-119">context</span><span class="sxs-lookup"><span data-stu-id="51408-119">context</span></span>](office.context.md) | <span data-ttu-id="51408-120">Composition</span><span class="sxs-lookup"><span data-stu-id="51408-120">Compose</span></span><br><span data-ttu-id="51408-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="51408-121">Read</span></span> | [<span data-ttu-id="51408-122">Context</span><span class="sxs-lookup"><span data-stu-id="51408-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="51408-123">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="51408-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="51408-124">Enumerations</span></span>

| <span data-ttu-id="51408-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="51408-125">Enumeration</span></span> | <span data-ttu-id="51408-126">Modes</span><span class="sxs-lookup"><span data-stu-id="51408-126">Modes</span></span> | <span data-ttu-id="51408-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="51408-127">Return type</span></span> | <span data-ttu-id="51408-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="51408-128">Minimum</span></span><br><span data-ttu-id="51408-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="51408-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="51408-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="51408-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="51408-131">Composition</span><span class="sxs-lookup"><span data-stu-id="51408-131">Compose</span></span><br><span data-ttu-id="51408-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="51408-132">Read</span></span> | <span data-ttu-id="51408-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-133">String</span></span> | [<span data-ttu-id="51408-134">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="51408-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="51408-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="51408-136">Composition</span><span class="sxs-lookup"><span data-stu-id="51408-136">Compose</span></span><br><span data-ttu-id="51408-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="51408-137">Read</span></span> | <span data-ttu-id="51408-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-138">String</span></span> | [<span data-ttu-id="51408-139">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="51408-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="51408-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="51408-141">Composition</span><span class="sxs-lookup"><span data-stu-id="51408-141">Compose</span></span><br><span data-ttu-id="51408-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="51408-142">Read</span></span> | <span data-ttu-id="51408-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-143">String</span></span> | [<span data-ttu-id="51408-144">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="51408-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="51408-145">Namespaces</span></span>

<span data-ttu-id="51408-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="51408-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="51408-147">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="51408-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="51408-148">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="51408-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="51408-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="51408-150">Type</span><span class="sxs-lookup"><span data-stu-id="51408-150">Type</span></span>

*   <span data-ttu-id="51408-151">String</span><span class="sxs-lookup"><span data-stu-id="51408-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="51408-152">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="51408-152">Properties:</span></span>

|<span data-ttu-id="51408-153">Nom</span><span class="sxs-lookup"><span data-stu-id="51408-153">Name</span></span>| <span data-ttu-id="51408-154">Type</span><span class="sxs-lookup"><span data-stu-id="51408-154">Type</span></span>| <span data-ttu-id="51408-155">Description</span><span class="sxs-lookup"><span data-stu-id="51408-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="51408-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-156">String</span></span>|<span data-ttu-id="51408-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="51408-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="51408-158">String</span><span class="sxs-lookup"><span data-stu-id="51408-158">String</span></span>|<span data-ttu-id="51408-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="51408-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="51408-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="51408-160">Requirements</span></span>

|<span data-ttu-id="51408-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="51408-161">Requirement</span></span>| <span data-ttu-id="51408-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="51408-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="51408-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="51408-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="51408-164">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-164">1.1</span></span>|
|[<span data-ttu-id="51408-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="51408-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="51408-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="51408-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="51408-167">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-167">CoercionType: String</span></span>

<span data-ttu-id="51408-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="51408-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="51408-169">Type</span><span class="sxs-lookup"><span data-stu-id="51408-169">Type</span></span>

*   <span data-ttu-id="51408-170">String</span><span class="sxs-lookup"><span data-stu-id="51408-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="51408-171">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="51408-171">Properties:</span></span>

|<span data-ttu-id="51408-172">Nom</span><span class="sxs-lookup"><span data-stu-id="51408-172">Name</span></span>| <span data-ttu-id="51408-173">Type</span><span class="sxs-lookup"><span data-stu-id="51408-173">Type</span></span>| <span data-ttu-id="51408-174">Description</span><span class="sxs-lookup"><span data-stu-id="51408-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="51408-175">Chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-175">String</span></span>|<span data-ttu-id="51408-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="51408-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="51408-177">String</span><span class="sxs-lookup"><span data-stu-id="51408-177">String</span></span>|<span data-ttu-id="51408-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="51408-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="51408-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="51408-179">Requirements</span></span>

|<span data-ttu-id="51408-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="51408-180">Requirement</span></span>| <span data-ttu-id="51408-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="51408-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="51408-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="51408-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="51408-183">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-183">1.1</span></span>|
|[<span data-ttu-id="51408-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="51408-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="51408-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="51408-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="51408-186">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-186">SourceProperty: String</span></span>

<span data-ttu-id="51408-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="51408-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="51408-188">Type</span><span class="sxs-lookup"><span data-stu-id="51408-188">Type</span></span>

*   <span data-ttu-id="51408-189">String</span><span class="sxs-lookup"><span data-stu-id="51408-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="51408-190">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="51408-190">Properties:</span></span>

|<span data-ttu-id="51408-191">Nom</span><span class="sxs-lookup"><span data-stu-id="51408-191">Name</span></span>| <span data-ttu-id="51408-192">Type</span><span class="sxs-lookup"><span data-stu-id="51408-192">Type</span></span>| <span data-ttu-id="51408-193">Description</span><span class="sxs-lookup"><span data-stu-id="51408-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="51408-194">Chaîne</span><span class="sxs-lookup"><span data-stu-id="51408-194">String</span></span>|<span data-ttu-id="51408-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="51408-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="51408-196">String</span><span class="sxs-lookup"><span data-stu-id="51408-196">String</span></span>|<span data-ttu-id="51408-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="51408-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="51408-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="51408-198">Requirements</span></span>

|<span data-ttu-id="51408-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="51408-199">Requirement</span></span>| <span data-ttu-id="51408-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="51408-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="51408-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="51408-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="51408-202">1.1</span><span class="sxs-lookup"><span data-stu-id="51408-202">1.1</span></span>|
|[<span data-ttu-id="51408-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="51408-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="51408-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="51408-204">Compose or Read</span></span>|
