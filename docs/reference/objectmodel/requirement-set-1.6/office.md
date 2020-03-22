---
title: Espace de noms Office-ensemble de conditions requises 1,6
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: dc7f62cc3f01e56f6c05b6cf40a4b73e87aea5e4
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891312"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="dc8ad-103">Office (boîte aux lettres requise définie sur 1,6)</span><span class="sxs-lookup"><span data-stu-id="dc8ad-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="dc8ad-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="dc8ad-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dc8ad-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc8ad-106">Requirements</span></span>

|<span data-ttu-id="dc8ad-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc8ad-107">Requirement</span></span>| <span data-ttu-id="dc8ad-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc8ad-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc8ad-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc8ad-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc8ad-110">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-110">1.1</span></span>|
|[<span data-ttu-id="dc8ad-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc8ad-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc8ad-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dc8ad-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="dc8ad-113">Properties</span></span>

| <span data-ttu-id="dc8ad-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="dc8ad-114">Property</span></span> | <span data-ttu-id="dc8ad-115">Modes</span><span class="sxs-lookup"><span data-stu-id="dc8ad-115">Modes</span></span> | <span data-ttu-id="dc8ad-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dc8ad-116">Return type</span></span> | <span data-ttu-id="dc8ad-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="dc8ad-117">Minimum</span></span><br><span data-ttu-id="dc8ad-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc8ad-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dc8ad-119">context</span><span class="sxs-lookup"><span data-stu-id="dc8ad-119">context</span></span>](office.context.md) | <span data-ttu-id="dc8ad-120">Composition</span><span class="sxs-lookup"><span data-stu-id="dc8ad-120">Compose</span></span><br><span data-ttu-id="dc8ad-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-121">Read</span></span> | [<span data-ttu-id="dc8ad-122">Context</span><span class="sxs-lookup"><span data-stu-id="dc8ad-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="dc8ad-123">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="dc8ad-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="dc8ad-124">Enumerations</span></span>

| <span data-ttu-id="dc8ad-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="dc8ad-125">Enumeration</span></span> | <span data-ttu-id="dc8ad-126">Modes</span><span class="sxs-lookup"><span data-stu-id="dc8ad-126">Modes</span></span> | <span data-ttu-id="dc8ad-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dc8ad-127">Return type</span></span> | <span data-ttu-id="dc8ad-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="dc8ad-128">Minimum</span></span><br><span data-ttu-id="dc8ad-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc8ad-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dc8ad-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dc8ad-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dc8ad-131">Composition</span><span class="sxs-lookup"><span data-stu-id="dc8ad-131">Compose</span></span><br><span data-ttu-id="dc8ad-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-132">Read</span></span> | <span data-ttu-id="dc8ad-133">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-133">String</span></span> | [<span data-ttu-id="dc8ad-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dc8ad-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dc8ad-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dc8ad-136">Composition</span><span class="sxs-lookup"><span data-stu-id="dc8ad-136">Compose</span></span><br><span data-ttu-id="dc8ad-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-137">Read</span></span> | <span data-ttu-id="dc8ad-138">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-138">String</span></span> | [<span data-ttu-id="dc8ad-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dc8ad-140">EventType</span><span class="sxs-lookup"><span data-stu-id="dc8ad-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="dc8ad-141">Composition</span><span class="sxs-lookup"><span data-stu-id="dc8ad-141">Compose</span></span><br><span data-ttu-id="dc8ad-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-142">Read</span></span> | <span data-ttu-id="dc8ad-143">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-143">String</span></span> | [<span data-ttu-id="dc8ad-144">1,5</span><span class="sxs-lookup"><span data-stu-id="dc8ad-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dc8ad-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dc8ad-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dc8ad-146">Composition</span><span class="sxs-lookup"><span data-stu-id="dc8ad-146">Compose</span></span><br><span data-ttu-id="dc8ad-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-147">Read</span></span> | <span data-ttu-id="dc8ad-148">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-148">String</span></span> | [<span data-ttu-id="dc8ad-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="dc8ad-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="dc8ad-150">Namespaces</span></span>

<span data-ttu-id="dc8ad-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="dc8ad-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="dc8ad-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="dc8ad-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="dc8ad-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="dc8ad-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dc8ad-155">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-155">Type</span></span>

*   <span data-ttu-id="dc8ad-156">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc8ad-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dc8ad-157">Properties:</span></span>

|<span data-ttu-id="dc8ad-158">Nom</span><span class="sxs-lookup"><span data-stu-id="dc8ad-158">Name</span></span>| <span data-ttu-id="dc8ad-159">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-159">Type</span></span>| <span data-ttu-id="dc8ad-160">Description</span><span class="sxs-lookup"><span data-stu-id="dc8ad-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dc8ad-161">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-161">String</span></span>|<span data-ttu-id="dc8ad-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dc8ad-163">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-163">String</span></span>|<span data-ttu-id="dc8ad-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc8ad-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc8ad-165">Requirements</span></span>

|<span data-ttu-id="dc8ad-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc8ad-166">Requirement</span></span>| <span data-ttu-id="dc8ad-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc8ad-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc8ad-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc8ad-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc8ad-169">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-169">1.1</span></span>|
|[<span data-ttu-id="dc8ad-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc8ad-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc8ad-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="dc8ad-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="dc8ad-172">CoercionType: String</span></span>

<span data-ttu-id="dc8ad-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc8ad-174">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-174">Type</span></span>

*   <span data-ttu-id="dc8ad-175">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc8ad-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dc8ad-176">Properties:</span></span>

|<span data-ttu-id="dc8ad-177">Nom</span><span class="sxs-lookup"><span data-stu-id="dc8ad-177">Name</span></span>| <span data-ttu-id="dc8ad-178">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-178">Type</span></span>| <span data-ttu-id="dc8ad-179">Description</span><span class="sxs-lookup"><span data-stu-id="dc8ad-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dc8ad-180">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-180">String</span></span>|<span data-ttu-id="dc8ad-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dc8ad-182">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-182">String</span></span>|<span data-ttu-id="dc8ad-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc8ad-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc8ad-184">Requirements</span></span>

|<span data-ttu-id="dc8ad-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc8ad-185">Requirement</span></span>| <span data-ttu-id="dc8ad-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc8ad-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc8ad-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc8ad-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc8ad-188">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-188">1.1</span></span>|
|[<span data-ttu-id="dc8ad-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc8ad-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc8ad-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="dc8ad-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="dc8ad-191">EventType: String</span></span>

<span data-ttu-id="dc8ad-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="dc8ad-193">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-193">Type</span></span>

*   <span data-ttu-id="dc8ad-194">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc8ad-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dc8ad-195">Properties:</span></span>

| <span data-ttu-id="dc8ad-196">Nom</span><span class="sxs-lookup"><span data-stu-id="dc8ad-196">Name</span></span> | <span data-ttu-id="dc8ad-197">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-197">Type</span></span> | <span data-ttu-id="dc8ad-198">Description</span><span class="sxs-lookup"><span data-stu-id="dc8ad-198">Description</span></span> | <span data-ttu-id="dc8ad-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="dc8ad-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="dc8ad-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dc8ad-200">String</span></span> | <span data-ttu-id="dc8ad-201">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="dc8ad-202">1,5</span><span class="sxs-lookup"><span data-stu-id="dc8ad-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dc8ad-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc8ad-203">Requirements</span></span>

|<span data-ttu-id="dc8ad-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc8ad-204">Requirement</span></span>| <span data-ttu-id="dc8ad-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc8ad-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc8ad-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc8ad-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc8ad-207">1,5</span><span class="sxs-lookup"><span data-stu-id="dc8ad-207">1.5</span></span> |
|[<span data-ttu-id="dc8ad-208">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc8ad-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc8ad-209">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="dc8ad-210">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="dc8ad-210">SourceProperty: String</span></span>

<span data-ttu-id="dc8ad-211">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dc8ad-212">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-212">Type</span></span>

*   <span data-ttu-id="dc8ad-213">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dc8ad-214">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dc8ad-214">Properties:</span></span>

|<span data-ttu-id="dc8ad-215">Nom</span><span class="sxs-lookup"><span data-stu-id="dc8ad-215">Name</span></span>| <span data-ttu-id="dc8ad-216">Type</span><span class="sxs-lookup"><span data-stu-id="dc8ad-216">Type</span></span>| <span data-ttu-id="dc8ad-217">Description</span><span class="sxs-lookup"><span data-stu-id="dc8ad-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dc8ad-218">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-218">String</span></span>|<span data-ttu-id="dc8ad-219">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dc8ad-220">String</span><span class="sxs-lookup"><span data-stu-id="dc8ad-220">String</span></span>|<span data-ttu-id="dc8ad-221">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="dc8ad-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dc8ad-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dc8ad-222">Requirements</span></span>

|<span data-ttu-id="dc8ad-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dc8ad-223">Requirement</span></span>| <span data-ttu-id="dc8ad-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="dc8ad-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="dc8ad-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dc8ad-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dc8ad-226">1.1</span><span class="sxs-lookup"><span data-stu-id="dc8ad-226">1.1</span></span>|
|[<span data-ttu-id="dc8ad-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dc8ad-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dc8ad-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dc8ad-228">Compose or Read</span></span>|
