---
title: Espace de noms Office-ensemble de conditions requises 1,8
description: L’espace de noms Office fournit des interfaces partagées pour les compléments Office Outlook (ensemble de conditions requises 1,8)
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0bbe212b0b8e5dc1348cb5cdc03509c44a716d1a
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717501"
---
# <a name="office"></a><span data-ttu-id="f2d53-103">Office</span><span class="sxs-lookup"><span data-stu-id="f2d53-103">Office</span></span>

<span data-ttu-id="f2d53-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f2d53-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f2d53-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f2d53-106">Requirements</span></span>

|<span data-ttu-id="f2d53-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2d53-107">Requirement</span></span>| <span data-ttu-id="f2d53-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="f2d53-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2d53-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f2d53-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f2d53-110">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-110">1.1</span></span>|
|[<span data-ttu-id="f2d53-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f2d53-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f2d53-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f2d53-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="f2d53-113">Properties</span></span>

| <span data-ttu-id="f2d53-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="f2d53-114">Property</span></span> | <span data-ttu-id="f2d53-115">Modes</span><span class="sxs-lookup"><span data-stu-id="f2d53-115">Modes</span></span> | <span data-ttu-id="f2d53-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f2d53-116">Return type</span></span> | <span data-ttu-id="f2d53-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="f2d53-117">Minimum</span></span><br><span data-ttu-id="f2d53-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2d53-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f2d53-119">context</span><span class="sxs-lookup"><span data-stu-id="f2d53-119">context</span></span>](office.context.md) | <span data-ttu-id="f2d53-120">Composition</span><span class="sxs-lookup"><span data-stu-id="f2d53-120">Compose</span></span><br><span data-ttu-id="f2d53-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-121">Read</span></span> | [<span data-ttu-id="f2d53-122">Context</span><span class="sxs-lookup"><span data-stu-id="f2d53-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="f2d53-123">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f2d53-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="f2d53-124">Enumerations</span></span>

| <span data-ttu-id="f2d53-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="f2d53-125">Enumeration</span></span> | <span data-ttu-id="f2d53-126">Modes</span><span class="sxs-lookup"><span data-stu-id="f2d53-126">Modes</span></span> | <span data-ttu-id="f2d53-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f2d53-127">Return type</span></span> | <span data-ttu-id="f2d53-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="f2d53-128">Minimum</span></span><br><span data-ttu-id="f2d53-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2d53-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f2d53-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f2d53-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f2d53-131">Composition</span><span class="sxs-lookup"><span data-stu-id="f2d53-131">Compose</span></span><br><span data-ttu-id="f2d53-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-132">Read</span></span> | <span data-ttu-id="f2d53-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-133">String</span></span> | [<span data-ttu-id="f2d53-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f2d53-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f2d53-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f2d53-136">Composition</span><span class="sxs-lookup"><span data-stu-id="f2d53-136">Compose</span></span><br><span data-ttu-id="f2d53-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-137">Read</span></span> | <span data-ttu-id="f2d53-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-138">String</span></span> | [<span data-ttu-id="f2d53-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f2d53-140">EventType</span><span class="sxs-lookup"><span data-stu-id="f2d53-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f2d53-141">Composition</span><span class="sxs-lookup"><span data-stu-id="f2d53-141">Compose</span></span><br><span data-ttu-id="f2d53-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-142">Read</span></span> | <span data-ttu-id="f2d53-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-143">String</span></span> | [<span data-ttu-id="f2d53-144">1,5</span><span class="sxs-lookup"><span data-stu-id="f2d53-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f2d53-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f2d53-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f2d53-146">Composition</span><span class="sxs-lookup"><span data-stu-id="f2d53-146">Compose</span></span><br><span data-ttu-id="f2d53-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-147">Read</span></span> | <span data-ttu-id="f2d53-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-148">String</span></span> | [<span data-ttu-id="f2d53-149">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f2d53-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="f2d53-150">Namespaces</span></span>

<span data-ttu-id="f2d53-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f2d53-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f2d53-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="f2d53-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f2d53-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="f2d53-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f2d53-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f2d53-155">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-155">Type</span></span>

*   <span data-ttu-id="f2d53-156">String</span><span class="sxs-lookup"><span data-stu-id="f2d53-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2d53-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f2d53-157">Properties:</span></span>

|<span data-ttu-id="f2d53-158">Nom</span><span class="sxs-lookup"><span data-stu-id="f2d53-158">Name</span></span>| <span data-ttu-id="f2d53-159">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-159">Type</span></span>| <span data-ttu-id="f2d53-160">Description</span><span class="sxs-lookup"><span data-stu-id="f2d53-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f2d53-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-161">String</span></span>|<span data-ttu-id="f2d53-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="f2d53-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f2d53-163">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-163">String</span></span>|<span data-ttu-id="f2d53-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="f2d53-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2d53-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f2d53-165">Requirements</span></span>

|<span data-ttu-id="f2d53-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2d53-166">Requirement</span></span>| <span data-ttu-id="f2d53-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="f2d53-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2d53-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f2d53-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f2d53-169">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-169">1.1</span></span>|
|[<span data-ttu-id="f2d53-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f2d53-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f2d53-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f2d53-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-172">CoercionType: String</span></span>

<span data-ttu-id="f2d53-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f2d53-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f2d53-174">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-174">Type</span></span>

*   <span data-ttu-id="f2d53-175">String</span><span class="sxs-lookup"><span data-stu-id="f2d53-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2d53-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f2d53-176">Properties:</span></span>

|<span data-ttu-id="f2d53-177">Nom</span><span class="sxs-lookup"><span data-stu-id="f2d53-177">Name</span></span>| <span data-ttu-id="f2d53-178">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-178">Type</span></span>| <span data-ttu-id="f2d53-179">Description</span><span class="sxs-lookup"><span data-stu-id="f2d53-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f2d53-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-180">String</span></span>|<span data-ttu-id="f2d53-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f2d53-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f2d53-182">String</span><span class="sxs-lookup"><span data-stu-id="f2d53-182">String</span></span>|<span data-ttu-id="f2d53-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="f2d53-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2d53-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f2d53-184">Requirements</span></span>

|<span data-ttu-id="f2d53-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2d53-185">Requirement</span></span>| <span data-ttu-id="f2d53-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="f2d53-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2d53-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f2d53-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f2d53-188">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-188">1.1</span></span>|
|[<span data-ttu-id="f2d53-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f2d53-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f2d53-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f2d53-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-191">EventType: String</span></span>

<span data-ttu-id="f2d53-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="f2d53-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f2d53-193">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-193">Type</span></span>

*   <span data-ttu-id="f2d53-194">String</span><span class="sxs-lookup"><span data-stu-id="f2d53-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2d53-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f2d53-195">Properties:</span></span>

| <span data-ttu-id="f2d53-196">Nom</span><span class="sxs-lookup"><span data-stu-id="f2d53-196">Name</span></span> | <span data-ttu-id="f2d53-197">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-197">Type</span></span> | <span data-ttu-id="f2d53-198">Description</span><span class="sxs-lookup"><span data-stu-id="f2d53-198">Description</span></span> | <span data-ttu-id="f2d53-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="f2d53-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="f2d53-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-200">String</span></span> | <span data-ttu-id="f2d53-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="f2d53-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f2d53-202">1.7</span><span class="sxs-lookup"><span data-stu-id="f2d53-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="f2d53-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-203">String</span></span> | <span data-ttu-id="f2d53-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="f2d53-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="f2d53-205">1.8</span><span class="sxs-lookup"><span data-stu-id="f2d53-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="f2d53-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-206">String</span></span> | <span data-ttu-id="f2d53-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="f2d53-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="f2d53-208">1.8</span><span class="sxs-lookup"><span data-stu-id="f2d53-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="f2d53-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-209">String</span></span> | <span data-ttu-id="f2d53-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="f2d53-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f2d53-211">1,5</span><span class="sxs-lookup"><span data-stu-id="f2d53-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f2d53-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-212">String</span></span> | <span data-ttu-id="f2d53-213">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="f2d53-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f2d53-214">1.7</span><span class="sxs-lookup"><span data-stu-id="f2d53-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f2d53-215">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-215">String</span></span> | <span data-ttu-id="f2d53-216">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="f2d53-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f2d53-217">1.7</span><span class="sxs-lookup"><span data-stu-id="f2d53-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f2d53-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f2d53-218">Requirements</span></span>

|<span data-ttu-id="f2d53-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2d53-219">Requirement</span></span>| <span data-ttu-id="f2d53-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="f2d53-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2d53-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f2d53-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f2d53-222">1,5</span><span class="sxs-lookup"><span data-stu-id="f2d53-222">1.5</span></span> |
|[<span data-ttu-id="f2d53-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f2d53-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f2d53-224">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f2d53-225">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-225">SourceProperty: String</span></span>

<span data-ttu-id="f2d53-226">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f2d53-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f2d53-227">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-227">Type</span></span>

*   <span data-ttu-id="f2d53-228">String</span><span class="sxs-lookup"><span data-stu-id="f2d53-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2d53-229">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f2d53-229">Properties:</span></span>

|<span data-ttu-id="f2d53-230">Nom</span><span class="sxs-lookup"><span data-stu-id="f2d53-230">Name</span></span>| <span data-ttu-id="f2d53-231">Type</span><span class="sxs-lookup"><span data-stu-id="f2d53-231">Type</span></span>| <span data-ttu-id="f2d53-232">Description</span><span class="sxs-lookup"><span data-stu-id="f2d53-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f2d53-233">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f2d53-233">String</span></span>|<span data-ttu-id="f2d53-234">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f2d53-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f2d53-235">String</span><span class="sxs-lookup"><span data-stu-id="f2d53-235">String</span></span>|<span data-ttu-id="f2d53-236">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="f2d53-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2d53-237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f2d53-237">Requirements</span></span>

|<span data-ttu-id="f2d53-238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f2d53-238">Requirement</span></span>| <span data-ttu-id="f2d53-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="f2d53-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2d53-240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f2d53-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f2d53-241">1.1</span><span class="sxs-lookup"><span data-stu-id="f2d53-241">1.1</span></span>|
|[<span data-ttu-id="f2d53-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f2d53-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f2d53-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f2d53-243">Compose or Read</span></span>|
