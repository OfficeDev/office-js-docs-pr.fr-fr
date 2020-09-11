---
title: Espace de noms Office-ensemble de conditions requises 1,8
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,8.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: e0580cd1bb327c8673c46d3d0292aec9f2f1c971
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431520"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="8fff0-103">Office (boîte aux lettres requise définie sur 1,8)</span><span class="sxs-lookup"><span data-stu-id="8fff0-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="8fff0-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="8fff0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8fff0-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8fff0-106">Requirements</span></span>

|<span data-ttu-id="8fff0-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8fff0-107">Requirement</span></span>| <span data-ttu-id="8fff0-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="8fff0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fff0-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8fff0-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fff0-110">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-110">1.1</span></span>|
|[<span data-ttu-id="8fff0-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8fff0-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fff0-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8fff0-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="8fff0-113">Properties</span></span>

| <span data-ttu-id="8fff0-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="8fff0-114">Property</span></span> | <span data-ttu-id="8fff0-115">Modes</span><span class="sxs-lookup"><span data-stu-id="8fff0-115">Modes</span></span> | <span data-ttu-id="8fff0-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="8fff0-116">Return type</span></span> | <span data-ttu-id="8fff0-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="8fff0-117">Minimum</span></span><br><span data-ttu-id="8fff0-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="8fff0-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8fff0-119">context</span><span class="sxs-lookup"><span data-stu-id="8fff0-119">context</span></span>](office.context.md) | <span data-ttu-id="8fff0-120">Composition</span><span class="sxs-lookup"><span data-stu-id="8fff0-120">Compose</span></span><br><span data-ttu-id="8fff0-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-121">Read</span></span> | [<span data-ttu-id="8fff0-122">Context</span><span class="sxs-lookup"><span data-stu-id="8fff0-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="8fff0-123">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8fff0-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="8fff0-124">Enumerations</span></span>

| <span data-ttu-id="8fff0-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="8fff0-125">Enumeration</span></span> | <span data-ttu-id="8fff0-126">Modes</span><span class="sxs-lookup"><span data-stu-id="8fff0-126">Modes</span></span> | <span data-ttu-id="8fff0-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="8fff0-127">Return type</span></span> | <span data-ttu-id="8fff0-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="8fff0-128">Minimum</span></span><br><span data-ttu-id="8fff0-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="8fff0-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8fff0-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8fff0-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8fff0-131">Composition</span><span class="sxs-lookup"><span data-stu-id="8fff0-131">Compose</span></span><br><span data-ttu-id="8fff0-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-132">Read</span></span> | <span data-ttu-id="8fff0-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-133">String</span></span> | [<span data-ttu-id="8fff0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fff0-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8fff0-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8fff0-136">Composition</span><span class="sxs-lookup"><span data-stu-id="8fff0-136">Compose</span></span><br><span data-ttu-id="8fff0-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-137">Read</span></span> | <span data-ttu-id="8fff0-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-138">String</span></span> | [<span data-ttu-id="8fff0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8fff0-140">EventType</span><span class="sxs-lookup"><span data-stu-id="8fff0-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8fff0-141">Composition</span><span class="sxs-lookup"><span data-stu-id="8fff0-141">Compose</span></span><br><span data-ttu-id="8fff0-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-142">Read</span></span> | <span data-ttu-id="8fff0-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-143">String</span></span> | [<span data-ttu-id="8fff0-144">1,5</span><span class="sxs-lookup"><span data-stu-id="8fff0-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8fff0-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8fff0-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8fff0-146">Composition</span><span class="sxs-lookup"><span data-stu-id="8fff0-146">Compose</span></span><br><span data-ttu-id="8fff0-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-147">Read</span></span> | <span data-ttu-id="8fff0-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-148">String</span></span> | [<span data-ttu-id="8fff0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8fff0-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="8fff0-150">Namespaces</span></span>

<span data-ttu-id="8fff0-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="8fff0-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8fff0-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="8fff0-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8fff0-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="8fff0-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8fff0-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8fff0-155">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-155">Type</span></span>

*   <span data-ttu-id="8fff0-156">String</span><span class="sxs-lookup"><span data-stu-id="8fff0-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8fff0-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8fff0-157">Properties:</span></span>

|<span data-ttu-id="8fff0-158">Nom</span><span class="sxs-lookup"><span data-stu-id="8fff0-158">Name</span></span>| <span data-ttu-id="8fff0-159">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-159">Type</span></span>| <span data-ttu-id="8fff0-160">Description</span><span class="sxs-lookup"><span data-stu-id="8fff0-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8fff0-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-161">String</span></span>|<span data-ttu-id="8fff0-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="8fff0-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8fff0-163">String</span><span class="sxs-lookup"><span data-stu-id="8fff0-163">String</span></span>|<span data-ttu-id="8fff0-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="8fff0-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fff0-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8fff0-165">Requirements</span></span>

|<span data-ttu-id="8fff0-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8fff0-166">Requirement</span></span>| <span data-ttu-id="8fff0-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="8fff0-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fff0-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8fff0-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fff0-169">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-169">1.1</span></span>|
|[<span data-ttu-id="8fff0-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8fff0-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fff0-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8fff0-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-172">CoercionType: String</span></span>

<span data-ttu-id="8fff0-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8fff0-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8fff0-174">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-174">Type</span></span>

*   <span data-ttu-id="8fff0-175">String</span><span class="sxs-lookup"><span data-stu-id="8fff0-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8fff0-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8fff0-176">Properties:</span></span>

|<span data-ttu-id="8fff0-177">Nom</span><span class="sxs-lookup"><span data-stu-id="8fff0-177">Name</span></span>| <span data-ttu-id="8fff0-178">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-178">Type</span></span>| <span data-ttu-id="8fff0-179">Description</span><span class="sxs-lookup"><span data-stu-id="8fff0-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8fff0-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-180">String</span></span>|<span data-ttu-id="8fff0-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="8fff0-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8fff0-182">String</span><span class="sxs-lookup"><span data-stu-id="8fff0-182">String</span></span>|<span data-ttu-id="8fff0-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="8fff0-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fff0-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8fff0-184">Requirements</span></span>

|<span data-ttu-id="8fff0-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8fff0-185">Requirement</span></span>| <span data-ttu-id="8fff0-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="8fff0-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fff0-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8fff0-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fff0-188">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-188">1.1</span></span>|
|[<span data-ttu-id="8fff0-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8fff0-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fff0-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="8fff0-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-191">EventType: String</span></span>

<span data-ttu-id="8fff0-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="8fff0-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8fff0-193">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-193">Type</span></span>

*   <span data-ttu-id="8fff0-194">String</span><span class="sxs-lookup"><span data-stu-id="8fff0-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8fff0-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8fff0-195">Properties:</span></span>

| <span data-ttu-id="8fff0-196">Nom</span><span class="sxs-lookup"><span data-stu-id="8fff0-196">Name</span></span> | <span data-ttu-id="8fff0-197">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-197">Type</span></span> | <span data-ttu-id="8fff0-198">Description</span><span class="sxs-lookup"><span data-stu-id="8fff0-198">Description</span></span> | <span data-ttu-id="8fff0-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="8fff0-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="8fff0-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-200">String</span></span> | <span data-ttu-id="8fff0-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="8fff0-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="8fff0-202">1.7</span><span class="sxs-lookup"><span data-stu-id="8fff0-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="8fff0-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-203">String</span></span> | <span data-ttu-id="8fff0-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="8fff0-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="8fff0-205">1.8</span><span class="sxs-lookup"><span data-stu-id="8fff0-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="8fff0-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-206">String</span></span> | <span data-ttu-id="8fff0-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="8fff0-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="8fff0-208">1.8</span><span class="sxs-lookup"><span data-stu-id="8fff0-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="8fff0-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-209">String</span></span> | <span data-ttu-id="8fff0-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="8fff0-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="8fff0-211">1,5</span><span class="sxs-lookup"><span data-stu-id="8fff0-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="8fff0-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-212">String</span></span> | <span data-ttu-id="8fff0-213">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="8fff0-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="8fff0-214">1.7</span><span class="sxs-lookup"><span data-stu-id="8fff0-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="8fff0-215">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-215">String</span></span> | <span data-ttu-id="8fff0-216">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="8fff0-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="8fff0-217">1.7</span><span class="sxs-lookup"><span data-stu-id="8fff0-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8fff0-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8fff0-218">Requirements</span></span>

|<span data-ttu-id="8fff0-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8fff0-219">Requirement</span></span>| <span data-ttu-id="8fff0-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="8fff0-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fff0-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8fff0-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fff0-222">1,5</span><span class="sxs-lookup"><span data-stu-id="8fff0-222">1.5</span></span> |
|[<span data-ttu-id="8fff0-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8fff0-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fff0-224">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8fff0-225">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-225">SourceProperty: String</span></span>

<span data-ttu-id="8fff0-226">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8fff0-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8fff0-227">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-227">Type</span></span>

*   <span data-ttu-id="8fff0-228">String</span><span class="sxs-lookup"><span data-stu-id="8fff0-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8fff0-229">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8fff0-229">Properties:</span></span>

|<span data-ttu-id="8fff0-230">Nom</span><span class="sxs-lookup"><span data-stu-id="8fff0-230">Name</span></span>| <span data-ttu-id="8fff0-231">Type</span><span class="sxs-lookup"><span data-stu-id="8fff0-231">Type</span></span>| <span data-ttu-id="8fff0-232">Description</span><span class="sxs-lookup"><span data-stu-id="8fff0-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8fff0-233">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8fff0-233">String</span></span>|<span data-ttu-id="8fff0-234">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="8fff0-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8fff0-235">String</span><span class="sxs-lookup"><span data-stu-id="8fff0-235">String</span></span>|<span data-ttu-id="8fff0-236">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="8fff0-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8fff0-237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8fff0-237">Requirements</span></span>

|<span data-ttu-id="8fff0-238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8fff0-238">Requirement</span></span>| <span data-ttu-id="8fff0-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="8fff0-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="8fff0-240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8fff0-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8fff0-241">1.1</span><span class="sxs-lookup"><span data-stu-id="8fff0-241">1.1</span></span>|
|[<span data-ttu-id="8fff0-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8fff0-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8fff0-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8fff0-243">Compose or Read</span></span>|
