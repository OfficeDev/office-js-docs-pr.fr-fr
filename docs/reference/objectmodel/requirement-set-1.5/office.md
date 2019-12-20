---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 63dbb3ac10492ac6e2019353b8cb057227e4c1e6
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814751"
---
# <a name="office"></a><span data-ttu-id="53de7-102">Office</span><span class="sxs-lookup"><span data-stu-id="53de7-102">Office</span></span>

<span data-ttu-id="53de7-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="53de7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="53de7-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="53de7-105">Requirements</span></span>

|<span data-ttu-id="53de7-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="53de7-106">Requirement</span></span>| <span data-ttu-id="53de7-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="53de7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="53de7-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="53de7-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="53de7-109">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-109">1.1</span></span>|
|[<span data-ttu-id="53de7-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="53de7-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53de7-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="53de7-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="53de7-112">Properties</span></span>

| <span data-ttu-id="53de7-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="53de7-113">Property</span></span> | <span data-ttu-id="53de7-114">Modes</span><span class="sxs-lookup"><span data-stu-id="53de7-114">Modes</span></span> | <span data-ttu-id="53de7-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="53de7-115">Return type</span></span> | <span data-ttu-id="53de7-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="53de7-116">Minimum</span></span><br><span data-ttu-id="53de7-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="53de7-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="53de7-118">context</span><span class="sxs-lookup"><span data-stu-id="53de7-118">context</span></span>](office.context.md) | <span data-ttu-id="53de7-119">Composition</span><span class="sxs-lookup"><span data-stu-id="53de7-119">Compose</span></span><br><span data-ttu-id="53de7-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-120">Read</span></span> | [<span data-ttu-id="53de7-121">Context</span><span class="sxs-lookup"><span data-stu-id="53de7-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="53de7-122">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="53de7-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="53de7-123">Enumerations</span></span>

| <span data-ttu-id="53de7-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="53de7-124">Enumeration</span></span> | <span data-ttu-id="53de7-125">Modes</span><span class="sxs-lookup"><span data-stu-id="53de7-125">Modes</span></span> | <span data-ttu-id="53de7-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="53de7-126">Return type</span></span> | <span data-ttu-id="53de7-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="53de7-127">Minimum</span></span><br><span data-ttu-id="53de7-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="53de7-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="53de7-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="53de7-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="53de7-130">Composition</span><span class="sxs-lookup"><span data-stu-id="53de7-130">Compose</span></span><br><span data-ttu-id="53de7-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-131">Read</span></span> | <span data-ttu-id="53de7-132">String</span><span class="sxs-lookup"><span data-stu-id="53de7-132">String</span></span> | [<span data-ttu-id="53de7-133">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="53de7-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="53de7-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="53de7-135">Composition</span><span class="sxs-lookup"><span data-stu-id="53de7-135">Compose</span></span><br><span data-ttu-id="53de7-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-136">Read</span></span> | <span data-ttu-id="53de7-137">String</span><span class="sxs-lookup"><span data-stu-id="53de7-137">String</span></span> | [<span data-ttu-id="53de7-138">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="53de7-139">EventType</span><span class="sxs-lookup"><span data-stu-id="53de7-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="53de7-140">Composition</span><span class="sxs-lookup"><span data-stu-id="53de7-140">Compose</span></span><br><span data-ttu-id="53de7-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-141">Read</span></span> | <span data-ttu-id="53de7-142">String</span><span class="sxs-lookup"><span data-stu-id="53de7-142">String</span></span> | [<span data-ttu-id="53de7-143">1,5</span><span class="sxs-lookup"><span data-stu-id="53de7-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="53de7-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="53de7-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="53de7-145">Composition</span><span class="sxs-lookup"><span data-stu-id="53de7-145">Compose</span></span><br><span data-ttu-id="53de7-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-146">Read</span></span> | <span data-ttu-id="53de7-147">String</span><span class="sxs-lookup"><span data-stu-id="53de7-147">String</span></span> | [<span data-ttu-id="53de7-148">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="53de7-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="53de7-149">Namespaces</span></span>

<span data-ttu-id="53de7-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="53de7-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="53de7-151">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="53de7-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="53de7-152">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="53de7-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="53de7-153">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="53de7-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="53de7-154">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-154">Type</span></span>

*   <span data-ttu-id="53de7-155">String</span><span class="sxs-lookup"><span data-stu-id="53de7-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53de7-156">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="53de7-156">Properties:</span></span>

|<span data-ttu-id="53de7-157">Nom</span><span class="sxs-lookup"><span data-stu-id="53de7-157">Name</span></span>| <span data-ttu-id="53de7-158">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-158">Type</span></span>| <span data-ttu-id="53de7-159">Description</span><span class="sxs-lookup"><span data-stu-id="53de7-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="53de7-160">String</span><span class="sxs-lookup"><span data-stu-id="53de7-160">String</span></span>|<span data-ttu-id="53de7-161">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="53de7-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="53de7-162">String</span><span class="sxs-lookup"><span data-stu-id="53de7-162">String</span></span>|<span data-ttu-id="53de7-163">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="53de7-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53de7-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="53de7-164">Requirements</span></span>

|<span data-ttu-id="53de7-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="53de7-165">Requirement</span></span>| <span data-ttu-id="53de7-166">Valeur</span><span class="sxs-lookup"><span data-stu-id="53de7-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="53de7-167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="53de7-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="53de7-168">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-168">1.1</span></span>|
|[<span data-ttu-id="53de7-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="53de7-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53de7-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="53de7-171">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="53de7-171">CoercionType: String</span></span>

<span data-ttu-id="53de7-172">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="53de7-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="53de7-173">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-173">Type</span></span>

*   <span data-ttu-id="53de7-174">String</span><span class="sxs-lookup"><span data-stu-id="53de7-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53de7-175">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="53de7-175">Properties:</span></span>

|<span data-ttu-id="53de7-176">Nom</span><span class="sxs-lookup"><span data-stu-id="53de7-176">Name</span></span>| <span data-ttu-id="53de7-177">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-177">Type</span></span>| <span data-ttu-id="53de7-178">Description</span><span class="sxs-lookup"><span data-stu-id="53de7-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="53de7-179">String</span><span class="sxs-lookup"><span data-stu-id="53de7-179">String</span></span>|<span data-ttu-id="53de7-180">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="53de7-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="53de7-181">String</span><span class="sxs-lookup"><span data-stu-id="53de7-181">String</span></span>|<span data-ttu-id="53de7-182">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="53de7-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53de7-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="53de7-183">Requirements</span></span>

|<span data-ttu-id="53de7-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="53de7-184">Requirement</span></span>| <span data-ttu-id="53de7-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="53de7-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="53de7-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="53de7-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="53de7-187">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-187">1.1</span></span>|
|[<span data-ttu-id="53de7-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="53de7-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53de7-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="53de7-190">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="53de7-190">EventType: String</span></span>

<span data-ttu-id="53de7-191">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="53de7-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="53de7-192">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-192">Type</span></span>

*   <span data-ttu-id="53de7-193">String</span><span class="sxs-lookup"><span data-stu-id="53de7-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53de7-194">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="53de7-194">Properties:</span></span>

| <span data-ttu-id="53de7-195">Nom</span><span class="sxs-lookup"><span data-stu-id="53de7-195">Name</span></span> | <span data-ttu-id="53de7-196">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-196">Type</span></span> | <span data-ttu-id="53de7-197">Description</span><span class="sxs-lookup"><span data-stu-id="53de7-197">Description</span></span> | <span data-ttu-id="53de7-198">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="53de7-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="53de7-199">Chaîne</span><span class="sxs-lookup"><span data-stu-id="53de7-199">String</span></span> | <span data-ttu-id="53de7-200">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="53de7-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="53de7-201">1,5</span><span class="sxs-lookup"><span data-stu-id="53de7-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="53de7-202">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="53de7-202">Requirements</span></span>

|<span data-ttu-id="53de7-203">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="53de7-203">Requirement</span></span>| <span data-ttu-id="53de7-204">Valeur</span><span class="sxs-lookup"><span data-stu-id="53de7-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="53de7-205">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="53de7-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="53de7-206">1,5</span><span class="sxs-lookup"><span data-stu-id="53de7-206">1.5</span></span> |
|[<span data-ttu-id="53de7-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="53de7-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53de7-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="53de7-209">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="53de7-209">SourceProperty: String</span></span>

<span data-ttu-id="53de7-210">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="53de7-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="53de7-211">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-211">Type</span></span>

*   <span data-ttu-id="53de7-212">String</span><span class="sxs-lookup"><span data-stu-id="53de7-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="53de7-213">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="53de7-213">Properties:</span></span>

|<span data-ttu-id="53de7-214">Nom</span><span class="sxs-lookup"><span data-stu-id="53de7-214">Name</span></span>| <span data-ttu-id="53de7-215">Type</span><span class="sxs-lookup"><span data-stu-id="53de7-215">Type</span></span>| <span data-ttu-id="53de7-216">Description</span><span class="sxs-lookup"><span data-stu-id="53de7-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="53de7-217">String</span><span class="sxs-lookup"><span data-stu-id="53de7-217">String</span></span>|<span data-ttu-id="53de7-218">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="53de7-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="53de7-219">String</span><span class="sxs-lookup"><span data-stu-id="53de7-219">String</span></span>|<span data-ttu-id="53de7-220">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="53de7-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="53de7-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="53de7-221">Requirements</span></span>

|<span data-ttu-id="53de7-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="53de7-222">Requirement</span></span>| <span data-ttu-id="53de7-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="53de7-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="53de7-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="53de7-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="53de7-225">1.1</span><span class="sxs-lookup"><span data-stu-id="53de7-225">1.1</span></span>|
|[<span data-ttu-id="53de7-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="53de7-226">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="53de7-227">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="53de7-227">Compose or Read</span></span>|
