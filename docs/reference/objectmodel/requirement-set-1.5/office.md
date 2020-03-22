---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 848aa30c07b936c8454b2833d5dce3e1d15ee193
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891347"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="93636-103">Office (boîte aux lettres requise définie sur 1,5)</span><span class="sxs-lookup"><span data-stu-id="93636-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="93636-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="93636-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="93636-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="93636-106">Requirements</span></span>

|<span data-ttu-id="93636-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="93636-107">Requirement</span></span>| <span data-ttu-id="93636-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="93636-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="93636-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="93636-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93636-110">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-110">1.1</span></span>|
|[<span data-ttu-id="93636-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="93636-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93636-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="93636-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="93636-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="93636-113">Properties</span></span>

| <span data-ttu-id="93636-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="93636-114">Property</span></span> | <span data-ttu-id="93636-115">Modes</span><span class="sxs-lookup"><span data-stu-id="93636-115">Modes</span></span> | <span data-ttu-id="93636-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="93636-116">Return type</span></span> | <span data-ttu-id="93636-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="93636-117">Minimum</span></span><br><span data-ttu-id="93636-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="93636-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="93636-119">context</span><span class="sxs-lookup"><span data-stu-id="93636-119">context</span></span>](office.context.md) | <span data-ttu-id="93636-120">Composition</span><span class="sxs-lookup"><span data-stu-id="93636-120">Compose</span></span><br><span data-ttu-id="93636-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="93636-121">Read</span></span> | [<span data-ttu-id="93636-122">Context</span><span class="sxs-lookup"><span data-stu-id="93636-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="93636-123">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="93636-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="93636-124">Enumerations</span></span>

| <span data-ttu-id="93636-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="93636-125">Enumeration</span></span> | <span data-ttu-id="93636-126">Modes</span><span class="sxs-lookup"><span data-stu-id="93636-126">Modes</span></span> | <span data-ttu-id="93636-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="93636-127">Return type</span></span> | <span data-ttu-id="93636-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="93636-128">Minimum</span></span><br><span data-ttu-id="93636-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="93636-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="93636-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="93636-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="93636-131">Composition</span><span class="sxs-lookup"><span data-stu-id="93636-131">Compose</span></span><br><span data-ttu-id="93636-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="93636-132">Read</span></span> | <span data-ttu-id="93636-133">String</span><span class="sxs-lookup"><span data-stu-id="93636-133">String</span></span> | [<span data-ttu-id="93636-134">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93636-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="93636-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="93636-136">Composition</span><span class="sxs-lookup"><span data-stu-id="93636-136">Compose</span></span><br><span data-ttu-id="93636-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="93636-137">Read</span></span> | <span data-ttu-id="93636-138">String</span><span class="sxs-lookup"><span data-stu-id="93636-138">String</span></span> | [<span data-ttu-id="93636-139">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="93636-140">EventType</span><span class="sxs-lookup"><span data-stu-id="93636-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="93636-141">Composition</span><span class="sxs-lookup"><span data-stu-id="93636-141">Compose</span></span><br><span data-ttu-id="93636-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="93636-142">Read</span></span> | <span data-ttu-id="93636-143">String</span><span class="sxs-lookup"><span data-stu-id="93636-143">String</span></span> | [<span data-ttu-id="93636-144">1,5</span><span class="sxs-lookup"><span data-stu-id="93636-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="93636-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="93636-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="93636-146">Composition</span><span class="sxs-lookup"><span data-stu-id="93636-146">Compose</span></span><br><span data-ttu-id="93636-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="93636-147">Read</span></span> | <span data-ttu-id="93636-148">String</span><span class="sxs-lookup"><span data-stu-id="93636-148">String</span></span> | [<span data-ttu-id="93636-149">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="93636-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="93636-150">Namespaces</span></span>

<span data-ttu-id="93636-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="93636-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="93636-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="93636-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="93636-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="93636-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="93636-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="93636-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="93636-155">Type</span><span class="sxs-lookup"><span data-stu-id="93636-155">Type</span></span>

*   <span data-ttu-id="93636-156">String</span><span class="sxs-lookup"><span data-stu-id="93636-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93636-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="93636-157">Properties:</span></span>

|<span data-ttu-id="93636-158">Nom</span><span class="sxs-lookup"><span data-stu-id="93636-158">Name</span></span>| <span data-ttu-id="93636-159">Type</span><span class="sxs-lookup"><span data-stu-id="93636-159">Type</span></span>| <span data-ttu-id="93636-160">Description</span><span class="sxs-lookup"><span data-stu-id="93636-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="93636-161">String</span><span class="sxs-lookup"><span data-stu-id="93636-161">String</span></span>|<span data-ttu-id="93636-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="93636-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="93636-163">String</span><span class="sxs-lookup"><span data-stu-id="93636-163">String</span></span>|<span data-ttu-id="93636-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="93636-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93636-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="93636-165">Requirements</span></span>

|<span data-ttu-id="93636-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="93636-166">Requirement</span></span>| <span data-ttu-id="93636-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="93636-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="93636-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="93636-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93636-169">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-169">1.1</span></span>|
|[<span data-ttu-id="93636-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="93636-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93636-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="93636-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="93636-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="93636-172">CoercionType: String</span></span>

<span data-ttu-id="93636-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="93636-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93636-174">Type</span><span class="sxs-lookup"><span data-stu-id="93636-174">Type</span></span>

*   <span data-ttu-id="93636-175">String</span><span class="sxs-lookup"><span data-stu-id="93636-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93636-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="93636-176">Properties:</span></span>

|<span data-ttu-id="93636-177">Nom</span><span class="sxs-lookup"><span data-stu-id="93636-177">Name</span></span>| <span data-ttu-id="93636-178">Type</span><span class="sxs-lookup"><span data-stu-id="93636-178">Type</span></span>| <span data-ttu-id="93636-179">Description</span><span class="sxs-lookup"><span data-stu-id="93636-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="93636-180">String</span><span class="sxs-lookup"><span data-stu-id="93636-180">String</span></span>|<span data-ttu-id="93636-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="93636-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="93636-182">String</span><span class="sxs-lookup"><span data-stu-id="93636-182">String</span></span>|<span data-ttu-id="93636-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="93636-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93636-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="93636-184">Requirements</span></span>

|<span data-ttu-id="93636-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="93636-185">Requirement</span></span>| <span data-ttu-id="93636-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="93636-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="93636-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="93636-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93636-188">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-188">1.1</span></span>|
|[<span data-ttu-id="93636-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="93636-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93636-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="93636-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="93636-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="93636-191">EventType: String</span></span>

<span data-ttu-id="93636-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="93636-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="93636-193">Type</span><span class="sxs-lookup"><span data-stu-id="93636-193">Type</span></span>

*   <span data-ttu-id="93636-194">String</span><span class="sxs-lookup"><span data-stu-id="93636-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93636-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="93636-195">Properties:</span></span>

| <span data-ttu-id="93636-196">Nom</span><span class="sxs-lookup"><span data-stu-id="93636-196">Name</span></span> | <span data-ttu-id="93636-197">Type</span><span class="sxs-lookup"><span data-stu-id="93636-197">Type</span></span> | <span data-ttu-id="93636-198">Description</span><span class="sxs-lookup"><span data-stu-id="93636-198">Description</span></span> | <span data-ttu-id="93636-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="93636-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="93636-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="93636-200">String</span></span> | <span data-ttu-id="93636-201">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="93636-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="93636-202">1,5</span><span class="sxs-lookup"><span data-stu-id="93636-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="93636-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="93636-203">Requirements</span></span>

|<span data-ttu-id="93636-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="93636-204">Requirement</span></span>| <span data-ttu-id="93636-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="93636-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="93636-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="93636-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93636-207">1,5</span><span class="sxs-lookup"><span data-stu-id="93636-207">1.5</span></span> |
|[<span data-ttu-id="93636-208">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="93636-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93636-209">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="93636-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="93636-210">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="93636-210">SourceProperty: String</span></span>

<span data-ttu-id="93636-211">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="93636-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="93636-212">Type</span><span class="sxs-lookup"><span data-stu-id="93636-212">Type</span></span>

*   <span data-ttu-id="93636-213">String</span><span class="sxs-lookup"><span data-stu-id="93636-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="93636-214">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="93636-214">Properties:</span></span>

|<span data-ttu-id="93636-215">Nom</span><span class="sxs-lookup"><span data-stu-id="93636-215">Name</span></span>| <span data-ttu-id="93636-216">Type</span><span class="sxs-lookup"><span data-stu-id="93636-216">Type</span></span>| <span data-ttu-id="93636-217">Description</span><span class="sxs-lookup"><span data-stu-id="93636-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="93636-218">String</span><span class="sxs-lookup"><span data-stu-id="93636-218">String</span></span>|<span data-ttu-id="93636-219">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="93636-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="93636-220">String</span><span class="sxs-lookup"><span data-stu-id="93636-220">String</span></span>|<span data-ttu-id="93636-221">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="93636-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="93636-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="93636-222">Requirements</span></span>

|<span data-ttu-id="93636-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="93636-223">Requirement</span></span>| <span data-ttu-id="93636-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="93636-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="93636-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="93636-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="93636-226">1.1</span><span class="sxs-lookup"><span data-stu-id="93636-226">1.1</span></span>|
|[<span data-ttu-id="93636-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="93636-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="93636-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="93636-228">Compose or Read</span></span>|
