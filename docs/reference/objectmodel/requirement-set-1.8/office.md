---
title: Espace de noms Office-ensemble de conditions requises 1,8
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,8.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 773a12d2f2b6c2d164b94d0b6b6c2dd0def90a41
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891179"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="4f31b-103">Office (boîte aux lettres requise définie sur 1,8)</span><span class="sxs-lookup"><span data-stu-id="4f31b-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="4f31b-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4f31b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4f31b-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f31b-106">Requirements</span></span>

|<span data-ttu-id="4f31b-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f31b-107">Requirement</span></span>| <span data-ttu-id="4f31b-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f31b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f31b-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f31b-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4f31b-110">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-110">1.1</span></span>|
|[<span data-ttu-id="4f31b-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f31b-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4f31b-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="4f31b-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="4f31b-113">Properties</span></span>

| <span data-ttu-id="4f31b-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="4f31b-114">Property</span></span> | <span data-ttu-id="4f31b-115">Modes</span><span class="sxs-lookup"><span data-stu-id="4f31b-115">Modes</span></span> | <span data-ttu-id="4f31b-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="4f31b-116">Return type</span></span> | <span data-ttu-id="4f31b-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="4f31b-117">Minimum</span></span><br><span data-ttu-id="4f31b-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f31b-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4f31b-119">context</span><span class="sxs-lookup"><span data-stu-id="4f31b-119">context</span></span>](office.context.md) | <span data-ttu-id="4f31b-120">Composition</span><span class="sxs-lookup"><span data-stu-id="4f31b-120">Compose</span></span><br><span data-ttu-id="4f31b-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-121">Read</span></span> | [<span data-ttu-id="4f31b-122">Context</span><span class="sxs-lookup"><span data-stu-id="4f31b-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8) | [<span data-ttu-id="4f31b-123">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="4f31b-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="4f31b-124">Enumerations</span></span>

| <span data-ttu-id="4f31b-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="4f31b-125">Enumeration</span></span> | <span data-ttu-id="4f31b-126">Modes</span><span class="sxs-lookup"><span data-stu-id="4f31b-126">Modes</span></span> | <span data-ttu-id="4f31b-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="4f31b-127">Return type</span></span> | <span data-ttu-id="4f31b-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="4f31b-128">Minimum</span></span><br><span data-ttu-id="4f31b-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f31b-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="4f31b-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4f31b-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4f31b-131">Composition</span><span class="sxs-lookup"><span data-stu-id="4f31b-131">Compose</span></span><br><span data-ttu-id="4f31b-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-132">Read</span></span> | <span data-ttu-id="4f31b-133">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-133">String</span></span> | [<span data-ttu-id="4f31b-134">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4f31b-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4f31b-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4f31b-136">Composition</span><span class="sxs-lookup"><span data-stu-id="4f31b-136">Compose</span></span><br><span data-ttu-id="4f31b-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-137">Read</span></span> | <span data-ttu-id="4f31b-138">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-138">String</span></span> | [<span data-ttu-id="4f31b-139">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="4f31b-140">EventType</span><span class="sxs-lookup"><span data-stu-id="4f31b-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4f31b-141">Composition</span><span class="sxs-lookup"><span data-stu-id="4f31b-141">Compose</span></span><br><span data-ttu-id="4f31b-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-142">Read</span></span> | <span data-ttu-id="4f31b-143">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-143">String</span></span> | [<span data-ttu-id="4f31b-144">1,5</span><span class="sxs-lookup"><span data-stu-id="4f31b-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="4f31b-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4f31b-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4f31b-146">Composition</span><span class="sxs-lookup"><span data-stu-id="4f31b-146">Compose</span></span><br><span data-ttu-id="4f31b-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-147">Read</span></span> | <span data-ttu-id="4f31b-148">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-148">String</span></span> | [<span data-ttu-id="4f31b-149">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="4f31b-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4f31b-150">Namespaces</span></span>

<span data-ttu-id="4f31b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="4f31b-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="4f31b-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="4f31b-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4f31b-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="4f31b-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="4f31b-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4f31b-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4f31b-155">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-155">Type</span></span>

*   <span data-ttu-id="4f31b-156">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4f31b-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4f31b-157">Properties:</span></span>

|<span data-ttu-id="4f31b-158">Nom</span><span class="sxs-lookup"><span data-stu-id="4f31b-158">Name</span></span>| <span data-ttu-id="4f31b-159">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-159">Type</span></span>| <span data-ttu-id="4f31b-160">Description</span><span class="sxs-lookup"><span data-stu-id="4f31b-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4f31b-161">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-161">String</span></span>|<span data-ttu-id="4f31b-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="4f31b-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4f31b-163">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-163">String</span></span>|<span data-ttu-id="4f31b-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="4f31b-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f31b-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f31b-165">Requirements</span></span>

|<span data-ttu-id="4f31b-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f31b-166">Requirement</span></span>| <span data-ttu-id="4f31b-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f31b-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f31b-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f31b-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4f31b-169">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-169">1.1</span></span>|
|[<span data-ttu-id="4f31b-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f31b-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4f31b-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="4f31b-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="4f31b-172">CoercionType: String</span></span>

<span data-ttu-id="4f31b-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4f31b-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4f31b-174">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-174">Type</span></span>

*   <span data-ttu-id="4f31b-175">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4f31b-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4f31b-176">Properties:</span></span>

|<span data-ttu-id="4f31b-177">Nom</span><span class="sxs-lookup"><span data-stu-id="4f31b-177">Name</span></span>| <span data-ttu-id="4f31b-178">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-178">Type</span></span>| <span data-ttu-id="4f31b-179">Description</span><span class="sxs-lookup"><span data-stu-id="4f31b-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4f31b-180">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-180">String</span></span>|<span data-ttu-id="4f31b-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4f31b-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4f31b-182">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-182">String</span></span>|<span data-ttu-id="4f31b-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="4f31b-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f31b-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f31b-184">Requirements</span></span>

|<span data-ttu-id="4f31b-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f31b-185">Requirement</span></span>| <span data-ttu-id="4f31b-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f31b-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f31b-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f31b-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4f31b-188">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-188">1.1</span></span>|
|[<span data-ttu-id="4f31b-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f31b-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4f31b-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="4f31b-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="4f31b-191">EventType: String</span></span>

<span data-ttu-id="4f31b-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4f31b-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4f31b-193">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-193">Type</span></span>

*   <span data-ttu-id="4f31b-194">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4f31b-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4f31b-195">Properties:</span></span>

| <span data-ttu-id="4f31b-196">Nom</span><span class="sxs-lookup"><span data-stu-id="4f31b-196">Name</span></span> | <span data-ttu-id="4f31b-197">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-197">Type</span></span> | <span data-ttu-id="4f31b-198">Description</span><span class="sxs-lookup"><span data-stu-id="4f31b-198">Description</span></span> | <span data-ttu-id="4f31b-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="4f31b-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="4f31b-200">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-200">String</span></span> | <span data-ttu-id="4f31b-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="4f31b-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4f31b-202">1.7</span><span class="sxs-lookup"><span data-stu-id="4f31b-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="4f31b-203">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-203">String</span></span> | <span data-ttu-id="4f31b-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="4f31b-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="4f31b-205">1.8</span><span class="sxs-lookup"><span data-stu-id="4f31b-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="4f31b-206">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-206">String</span></span> | <span data-ttu-id="4f31b-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="4f31b-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="4f31b-208">1.8</span><span class="sxs-lookup"><span data-stu-id="4f31b-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="4f31b-209">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-209">String</span></span> | <span data-ttu-id="4f31b-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="4f31b-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="4f31b-211">1,5</span><span class="sxs-lookup"><span data-stu-id="4f31b-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4f31b-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4f31b-212">String</span></span> | <span data-ttu-id="4f31b-213">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="4f31b-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4f31b-214">1.7</span><span class="sxs-lookup"><span data-stu-id="4f31b-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4f31b-215">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4f31b-215">String</span></span> | <span data-ttu-id="4f31b-216">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="4f31b-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4f31b-217">1.7</span><span class="sxs-lookup"><span data-stu-id="4f31b-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4f31b-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f31b-218">Requirements</span></span>

|<span data-ttu-id="4f31b-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f31b-219">Requirement</span></span>| <span data-ttu-id="4f31b-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f31b-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f31b-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f31b-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4f31b-222">1,5</span><span class="sxs-lookup"><span data-stu-id="4f31b-222">1.5</span></span> |
|[<span data-ttu-id="4f31b-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f31b-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4f31b-224">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4f31b-225">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="4f31b-225">SourceProperty: String</span></span>

<span data-ttu-id="4f31b-226">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4f31b-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4f31b-227">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-227">Type</span></span>

*   <span data-ttu-id="4f31b-228">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4f31b-229">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4f31b-229">Properties:</span></span>

|<span data-ttu-id="4f31b-230">Nom</span><span class="sxs-lookup"><span data-stu-id="4f31b-230">Name</span></span>| <span data-ttu-id="4f31b-231">Type</span><span class="sxs-lookup"><span data-stu-id="4f31b-231">Type</span></span>| <span data-ttu-id="4f31b-232">Description</span><span class="sxs-lookup"><span data-stu-id="4f31b-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4f31b-233">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-233">String</span></span>|<span data-ttu-id="4f31b-234">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4f31b-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4f31b-235">String</span><span class="sxs-lookup"><span data-stu-id="4f31b-235">String</span></span>|<span data-ttu-id="4f31b-236">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4f31b-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4f31b-237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4f31b-237">Requirements</span></span>

|<span data-ttu-id="4f31b-238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4f31b-238">Requirement</span></span>| <span data-ttu-id="4f31b-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="4f31b-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="4f31b-240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4f31b-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="4f31b-241">1.1</span><span class="sxs-lookup"><span data-stu-id="4f31b-241">1.1</span></span>|
|[<span data-ttu-id="4f31b-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4f31b-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="4f31b-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4f31b-243">Compose or Read</span></span>|
