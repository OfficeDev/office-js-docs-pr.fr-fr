---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7cc8e6acc60c28b44ec7a2b91bb5e388b2618a31
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688851"
---
# <a name="office"></a><span data-ttu-id="dab97-102">Office</span><span class="sxs-lookup"><span data-stu-id="dab97-102">Office</span></span>

<span data-ttu-id="dab97-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="dab97-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dab97-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dab97-105">Requirements</span></span>

|<span data-ttu-id="dab97-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dab97-106">Requirement</span></span>| <span data-ttu-id="dab97-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="dab97-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="dab97-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dab97-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dab97-109">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-109">1.1</span></span>|
|[<span data-ttu-id="dab97-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dab97-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dab97-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dab97-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="dab97-112">Properties</span></span>

| <span data-ttu-id="dab97-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="dab97-113">Property</span></span> | <span data-ttu-id="dab97-114">Modes</span><span class="sxs-lookup"><span data-stu-id="dab97-114">Modes</span></span> | <span data-ttu-id="dab97-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dab97-115">Return type</span></span> | <span data-ttu-id="dab97-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="dab97-116">Minimum</span></span><br><span data-ttu-id="dab97-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dab97-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dab97-118">context</span><span class="sxs-lookup"><span data-stu-id="dab97-118">context</span></span>](office.context.md) | <span data-ttu-id="dab97-119">Composition</span><span class="sxs-lookup"><span data-stu-id="dab97-119">Compose</span></span><br><span data-ttu-id="dab97-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-120">Read</span></span> | [<span data-ttu-id="dab97-121">Context</span><span class="sxs-lookup"><span data-stu-id="dab97-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="dab97-122">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="dab97-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="dab97-123">Enumerations</span></span>

| <span data-ttu-id="dab97-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="dab97-124">Enumeration</span></span> | <span data-ttu-id="dab97-125">Modes</span><span class="sxs-lookup"><span data-stu-id="dab97-125">Modes</span></span> | <span data-ttu-id="dab97-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dab97-126">Return type</span></span> | <span data-ttu-id="dab97-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="dab97-127">Minimum</span></span><br><span data-ttu-id="dab97-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dab97-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dab97-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dab97-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dab97-130">Composition</span><span class="sxs-lookup"><span data-stu-id="dab97-130">Compose</span></span><br><span data-ttu-id="dab97-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-131">Read</span></span> | <span data-ttu-id="dab97-132">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-132">String</span></span> | [<span data-ttu-id="dab97-133">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dab97-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dab97-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dab97-135">Composition</span><span class="sxs-lookup"><span data-stu-id="dab97-135">Compose</span></span><br><span data-ttu-id="dab97-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-136">Read</span></span> | <span data-ttu-id="dab97-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-137">String</span></span> | [<span data-ttu-id="dab97-138">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dab97-139">EventType</span><span class="sxs-lookup"><span data-stu-id="dab97-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="dab97-140">Composition</span><span class="sxs-lookup"><span data-stu-id="dab97-140">Compose</span></span><br><span data-ttu-id="dab97-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-141">Read</span></span> | <span data-ttu-id="dab97-142">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-142">String</span></span> | [<span data-ttu-id="dab97-143">1,5</span><span class="sxs-lookup"><span data-stu-id="dab97-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dab97-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dab97-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dab97-145">Composition</span><span class="sxs-lookup"><span data-stu-id="dab97-145">Compose</span></span><br><span data-ttu-id="dab97-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-146">Read</span></span> | <span data-ttu-id="dab97-147">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-147">String</span></span> | [<span data-ttu-id="dab97-148">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="dab97-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="dab97-149">Namespaces</span></span>

<span data-ttu-id="dab97-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="dab97-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="dab97-151">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="dab97-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="dab97-152">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="dab97-153">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dab97-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dab97-154">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-154">Type</span></span>

*   <span data-ttu-id="dab97-155">String</span><span class="sxs-lookup"><span data-stu-id="dab97-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dab97-156">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dab97-156">Properties:</span></span>

|<span data-ttu-id="dab97-157">Nom</span><span class="sxs-lookup"><span data-stu-id="dab97-157">Name</span></span>| <span data-ttu-id="dab97-158">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-158">Type</span></span>| <span data-ttu-id="dab97-159">Description</span><span class="sxs-lookup"><span data-stu-id="dab97-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dab97-160">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-160">String</span></span>|<span data-ttu-id="dab97-161">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="dab97-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dab97-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-162">String</span></span>|<span data-ttu-id="dab97-163">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="dab97-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dab97-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dab97-164">Requirements</span></span>

|<span data-ttu-id="dab97-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dab97-165">Requirement</span></span>| <span data-ttu-id="dab97-166">Valeur</span><span class="sxs-lookup"><span data-stu-id="dab97-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="dab97-167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dab97-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dab97-168">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-168">1.1</span></span>|
|[<span data-ttu-id="dab97-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dab97-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dab97-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="dab97-171">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-171">CoercionType: String</span></span>

<span data-ttu-id="dab97-172">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="dab97-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dab97-173">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-173">Type</span></span>

*   <span data-ttu-id="dab97-174">String</span><span class="sxs-lookup"><span data-stu-id="dab97-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dab97-175">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dab97-175">Properties:</span></span>

|<span data-ttu-id="dab97-176">Nom</span><span class="sxs-lookup"><span data-stu-id="dab97-176">Name</span></span>| <span data-ttu-id="dab97-177">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-177">Type</span></span>| <span data-ttu-id="dab97-178">Description</span><span class="sxs-lookup"><span data-stu-id="dab97-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dab97-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-179">String</span></span>|<span data-ttu-id="dab97-180">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="dab97-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dab97-181">String</span><span class="sxs-lookup"><span data-stu-id="dab97-181">String</span></span>|<span data-ttu-id="dab97-182">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="dab97-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dab97-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dab97-183">Requirements</span></span>

|<span data-ttu-id="dab97-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dab97-184">Requirement</span></span>| <span data-ttu-id="dab97-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="dab97-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="dab97-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dab97-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dab97-187">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-187">1.1</span></span>|
|[<span data-ttu-id="dab97-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dab97-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dab97-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="dab97-190">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-190">EventType: String</span></span>

<span data-ttu-id="dab97-191">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="dab97-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="dab97-192">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-192">Type</span></span>

*   <span data-ttu-id="dab97-193">String</span><span class="sxs-lookup"><span data-stu-id="dab97-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dab97-194">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dab97-194">Properties:</span></span>

| <span data-ttu-id="dab97-195">Nom</span><span class="sxs-lookup"><span data-stu-id="dab97-195">Name</span></span> | <span data-ttu-id="dab97-196">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-196">Type</span></span> | <span data-ttu-id="dab97-197">Description</span><span class="sxs-lookup"><span data-stu-id="dab97-197">Description</span></span> | <span data-ttu-id="dab97-198">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="dab97-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="dab97-199">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-199">String</span></span> | <span data-ttu-id="dab97-200">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="dab97-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="dab97-201">1,5</span><span class="sxs-lookup"><span data-stu-id="dab97-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dab97-202">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dab97-202">Requirements</span></span>

|<span data-ttu-id="dab97-203">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dab97-203">Requirement</span></span>| <span data-ttu-id="dab97-204">Valeur</span><span class="sxs-lookup"><span data-stu-id="dab97-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="dab97-205">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dab97-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dab97-206">1,5</span><span class="sxs-lookup"><span data-stu-id="dab97-206">1.5</span></span> |
|[<span data-ttu-id="dab97-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dab97-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dab97-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="dab97-209">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-209">SourceProperty: String</span></span>

<span data-ttu-id="dab97-210">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="dab97-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dab97-211">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-211">Type</span></span>

*   <span data-ttu-id="dab97-212">String</span><span class="sxs-lookup"><span data-stu-id="dab97-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dab97-213">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dab97-213">Properties:</span></span>

|<span data-ttu-id="dab97-214">Nom</span><span class="sxs-lookup"><span data-stu-id="dab97-214">Name</span></span>| <span data-ttu-id="dab97-215">Type</span><span class="sxs-lookup"><span data-stu-id="dab97-215">Type</span></span>| <span data-ttu-id="dab97-216">Description</span><span class="sxs-lookup"><span data-stu-id="dab97-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dab97-217">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dab97-217">String</span></span>|<span data-ttu-id="dab97-218">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="dab97-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dab97-219">String</span><span class="sxs-lookup"><span data-stu-id="dab97-219">String</span></span>|<span data-ttu-id="dab97-220">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="dab97-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dab97-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dab97-221">Requirements</span></span>

|<span data-ttu-id="dab97-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dab97-222">Requirement</span></span>| <span data-ttu-id="dab97-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="dab97-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="dab97-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dab97-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dab97-225">1.1</span><span class="sxs-lookup"><span data-stu-id="dab97-225">1.1</span></span>|
|[<span data-ttu-id="dab97-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dab97-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dab97-227">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dab97-227">Compose or Read</span></span>|
