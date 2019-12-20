---
title: Espace de noms Office-ensemble de conditions requises 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b23afd7b84dcd18e120f6aea4bd4fb0952791f1c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814165"
---
# <a name="office"></a><span data-ttu-id="ac042-102">Office</span><span class="sxs-lookup"><span data-stu-id="ac042-102">Office</span></span>

<span data-ttu-id="ac042-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="ac042-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac042-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac042-105">Requirements</span></span>

|<span data-ttu-id="ac042-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac042-106">Requirement</span></span>| <span data-ttu-id="ac042-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac042-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac042-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac042-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac042-109">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-109">1.1</span></span>|
|[<span data-ttu-id="ac042-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac042-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac042-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ac042-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ac042-112">Properties</span></span>

| <span data-ttu-id="ac042-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="ac042-113">Property</span></span> | <span data-ttu-id="ac042-114">Modes</span><span class="sxs-lookup"><span data-stu-id="ac042-114">Modes</span></span> | <span data-ttu-id="ac042-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ac042-115">Return type</span></span> | <span data-ttu-id="ac042-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="ac042-116">Minimum</span></span><br><span data-ttu-id="ac042-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac042-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac042-118">context</span><span class="sxs-lookup"><span data-stu-id="ac042-118">context</span></span>](office.context.md) | <span data-ttu-id="ac042-119">Composition</span><span class="sxs-lookup"><span data-stu-id="ac042-119">Compose</span></span><br><span data-ttu-id="ac042-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-120">Read</span></span> | [<span data-ttu-id="ac042-121">Context</span><span class="sxs-lookup"><span data-stu-id="ac042-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="ac042-122">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="ac042-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="ac042-123">Enumerations</span></span>

| <span data-ttu-id="ac042-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="ac042-124">Enumeration</span></span> | <span data-ttu-id="ac042-125">Modes</span><span class="sxs-lookup"><span data-stu-id="ac042-125">Modes</span></span> | <span data-ttu-id="ac042-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ac042-126">Return type</span></span> | <span data-ttu-id="ac042-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="ac042-127">Minimum</span></span><br><span data-ttu-id="ac042-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac042-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac042-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="ac042-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="ac042-130">Composition</span><span class="sxs-lookup"><span data-stu-id="ac042-130">Compose</span></span><br><span data-ttu-id="ac042-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-131">Read</span></span> | <span data-ttu-id="ac042-132">String</span><span class="sxs-lookup"><span data-stu-id="ac042-132">String</span></span> | [<span data-ttu-id="ac042-133">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac042-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="ac042-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="ac042-135">Composition</span><span class="sxs-lookup"><span data-stu-id="ac042-135">Compose</span></span><br><span data-ttu-id="ac042-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-136">Read</span></span> | <span data-ttu-id="ac042-137">String</span><span class="sxs-lookup"><span data-stu-id="ac042-137">String</span></span> | [<span data-ttu-id="ac042-138">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac042-139">EventType</span><span class="sxs-lookup"><span data-stu-id="ac042-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="ac042-140">Composition</span><span class="sxs-lookup"><span data-stu-id="ac042-140">Compose</span></span><br><span data-ttu-id="ac042-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-141">Read</span></span> | <span data-ttu-id="ac042-142">String</span><span class="sxs-lookup"><span data-stu-id="ac042-142">String</span></span> | [<span data-ttu-id="ac042-143">1,5</span><span class="sxs-lookup"><span data-stu-id="ac042-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ac042-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="ac042-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="ac042-145">Composition</span><span class="sxs-lookup"><span data-stu-id="ac042-145">Compose</span></span><br><span data-ttu-id="ac042-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-146">Read</span></span> | <span data-ttu-id="ac042-147">String</span><span class="sxs-lookup"><span data-stu-id="ac042-147">String</span></span> | [<span data-ttu-id="ac042-148">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="ac042-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="ac042-149">Namespaces</span></span>

<span data-ttu-id="ac042-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="ac042-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="ac042-151">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="ac042-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="ac042-152">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="ac042-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="ac042-153">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="ac042-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="ac042-154">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-154">Type</span></span>

*   <span data-ttu-id="ac042-155">String</span><span class="sxs-lookup"><span data-stu-id="ac042-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac042-156">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ac042-156">Properties:</span></span>

|<span data-ttu-id="ac042-157">Nom</span><span class="sxs-lookup"><span data-stu-id="ac042-157">Name</span></span>| <span data-ttu-id="ac042-158">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-158">Type</span></span>| <span data-ttu-id="ac042-159">Description</span><span class="sxs-lookup"><span data-stu-id="ac042-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="ac042-160">String</span><span class="sxs-lookup"><span data-stu-id="ac042-160">String</span></span>|<span data-ttu-id="ac042-161">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="ac042-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="ac042-162">String</span><span class="sxs-lookup"><span data-stu-id="ac042-162">String</span></span>|<span data-ttu-id="ac042-163">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="ac042-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac042-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac042-164">Requirements</span></span>

|<span data-ttu-id="ac042-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac042-165">Requirement</span></span>| <span data-ttu-id="ac042-166">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac042-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac042-167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac042-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac042-168">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-168">1.1</span></span>|
|[<span data-ttu-id="ac042-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac042-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac042-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="ac042-171">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="ac042-171">CoercionType: String</span></span>

<span data-ttu-id="ac042-172">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="ac042-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ac042-173">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-173">Type</span></span>

*   <span data-ttu-id="ac042-174">String</span><span class="sxs-lookup"><span data-stu-id="ac042-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac042-175">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ac042-175">Properties:</span></span>

|<span data-ttu-id="ac042-176">Nom</span><span class="sxs-lookup"><span data-stu-id="ac042-176">Name</span></span>| <span data-ttu-id="ac042-177">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-177">Type</span></span>| <span data-ttu-id="ac042-178">Description</span><span class="sxs-lookup"><span data-stu-id="ac042-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="ac042-179">String</span><span class="sxs-lookup"><span data-stu-id="ac042-179">String</span></span>|<span data-ttu-id="ac042-180">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="ac042-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="ac042-181">String</span><span class="sxs-lookup"><span data-stu-id="ac042-181">String</span></span>|<span data-ttu-id="ac042-182">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="ac042-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac042-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac042-183">Requirements</span></span>

|<span data-ttu-id="ac042-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac042-184">Requirement</span></span>| <span data-ttu-id="ac042-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac042-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac042-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac042-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac042-187">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-187">1.1</span></span>|
|[<span data-ttu-id="ac042-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac042-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac042-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="ac042-190">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="ac042-190">EventType: String</span></span>

<span data-ttu-id="ac042-191">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="ac042-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="ac042-192">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-192">Type</span></span>

*   <span data-ttu-id="ac042-193">String</span><span class="sxs-lookup"><span data-stu-id="ac042-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac042-194">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ac042-194">Properties:</span></span>

| <span data-ttu-id="ac042-195">Nom</span><span class="sxs-lookup"><span data-stu-id="ac042-195">Name</span></span> | <span data-ttu-id="ac042-196">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-196">Type</span></span> | <span data-ttu-id="ac042-197">Description</span><span class="sxs-lookup"><span data-stu-id="ac042-197">Description</span></span> | <span data-ttu-id="ac042-198">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="ac042-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="ac042-199">String</span><span class="sxs-lookup"><span data-stu-id="ac042-199">String</span></span> | <span data-ttu-id="ac042-200">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="ac042-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="ac042-201">1.7</span><span class="sxs-lookup"><span data-stu-id="ac042-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="ac042-202">String</span><span class="sxs-lookup"><span data-stu-id="ac042-202">String</span></span> | <span data-ttu-id="ac042-203">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="ac042-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="ac042-204">1.8</span><span class="sxs-lookup"><span data-stu-id="ac042-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="ac042-205">String</span><span class="sxs-lookup"><span data-stu-id="ac042-205">String</span></span> | <span data-ttu-id="ac042-206">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="ac042-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="ac042-207">1.8</span><span class="sxs-lookup"><span data-stu-id="ac042-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="ac042-208">String</span><span class="sxs-lookup"><span data-stu-id="ac042-208">String</span></span> | <span data-ttu-id="ac042-209">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="ac042-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="ac042-210">1,5</span><span class="sxs-lookup"><span data-stu-id="ac042-210">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="ac042-211">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ac042-211">String</span></span> | <span data-ttu-id="ac042-212">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="ac042-212">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="ac042-213">1.7</span><span class="sxs-lookup"><span data-stu-id="ac042-213">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="ac042-214">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ac042-214">String</span></span> | <span data-ttu-id="ac042-215">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="ac042-215">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="ac042-216">1.7</span><span class="sxs-lookup"><span data-stu-id="ac042-216">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="ac042-217">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac042-217">Requirements</span></span>

|<span data-ttu-id="ac042-218">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac042-218">Requirement</span></span>| <span data-ttu-id="ac042-219">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac042-219">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac042-220">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac042-220">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac042-221">1,5</span><span class="sxs-lookup"><span data-stu-id="ac042-221">1.5</span></span> |
|[<span data-ttu-id="ac042-222">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac042-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac042-223">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-223">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="ac042-224">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="ac042-224">SourceProperty: String</span></span>

<span data-ttu-id="ac042-225">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="ac042-225">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="ac042-226">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-226">Type</span></span>

*   <span data-ttu-id="ac042-227">String</span><span class="sxs-lookup"><span data-stu-id="ac042-227">String</span></span>

##### <a name="properties"></a><span data-ttu-id="ac042-228">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ac042-228">Properties:</span></span>

|<span data-ttu-id="ac042-229">Nom</span><span class="sxs-lookup"><span data-stu-id="ac042-229">Name</span></span>| <span data-ttu-id="ac042-230">Type</span><span class="sxs-lookup"><span data-stu-id="ac042-230">Type</span></span>| <span data-ttu-id="ac042-231">Description</span><span class="sxs-lookup"><span data-stu-id="ac042-231">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="ac042-232">String</span><span class="sxs-lookup"><span data-stu-id="ac042-232">String</span></span>|<span data-ttu-id="ac042-233">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="ac042-233">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="ac042-234">String</span><span class="sxs-lookup"><span data-stu-id="ac042-234">String</span></span>|<span data-ttu-id="ac042-235">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="ac042-235">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ac042-236">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac042-236">Requirements</span></span>

|<span data-ttu-id="ac042-237">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac042-237">Requirement</span></span>| <span data-ttu-id="ac042-238">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac042-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac042-239">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac042-239">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac042-240">1.1</span><span class="sxs-lookup"><span data-stu-id="ac042-240">1.1</span></span>|
|[<span data-ttu-id="ac042-241">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac042-241">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ac042-242">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac042-242">Compose or Read</span></span>|
