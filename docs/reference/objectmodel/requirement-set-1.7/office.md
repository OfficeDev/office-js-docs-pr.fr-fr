---
title: Espace de noms Office-ensemble de conditions requises 1,7
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: ed98cada1328c32caa79279981bd0ce555a17385
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431394"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="209f3-103">Office (boîte aux lettres requise définie sur 1,7)</span><span class="sxs-lookup"><span data-stu-id="209f3-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="209f3-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="209f3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="209f3-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="209f3-106">Requirements</span></span>

|<span data-ttu-id="209f3-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="209f3-107">Requirement</span></span>| <span data-ttu-id="209f3-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="209f3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="209f3-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="209f3-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="209f3-110">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-110">1.1</span></span>|
|[<span data-ttu-id="209f3-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="209f3-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="209f3-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="209f3-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="209f3-113">Properties</span></span>

| <span data-ttu-id="209f3-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="209f3-114">Property</span></span> | <span data-ttu-id="209f3-115">Modes</span><span class="sxs-lookup"><span data-stu-id="209f3-115">Modes</span></span> | <span data-ttu-id="209f3-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="209f3-116">Return type</span></span> | <span data-ttu-id="209f3-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="209f3-117">Minimum</span></span><br><span data-ttu-id="209f3-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="209f3-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="209f3-119">context</span><span class="sxs-lookup"><span data-stu-id="209f3-119">context</span></span>](office.context.md) | <span data-ttu-id="209f3-120">Composition</span><span class="sxs-lookup"><span data-stu-id="209f3-120">Compose</span></span><br><span data-ttu-id="209f3-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-121">Read</span></span> | [<span data-ttu-id="209f3-122">Context</span><span class="sxs-lookup"><span data-stu-id="209f3-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="209f3-123">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="209f3-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="209f3-124">Enumerations</span></span>

| <span data-ttu-id="209f3-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="209f3-125">Enumeration</span></span> | <span data-ttu-id="209f3-126">Modes</span><span class="sxs-lookup"><span data-stu-id="209f3-126">Modes</span></span> | <span data-ttu-id="209f3-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="209f3-127">Return type</span></span> | <span data-ttu-id="209f3-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="209f3-128">Minimum</span></span><br><span data-ttu-id="209f3-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="209f3-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="209f3-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="209f3-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="209f3-131">Composition</span><span class="sxs-lookup"><span data-stu-id="209f3-131">Compose</span></span><br><span data-ttu-id="209f3-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-132">Read</span></span> | <span data-ttu-id="209f3-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-133">String</span></span> | [<span data-ttu-id="209f3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="209f3-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="209f3-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="209f3-136">Composition</span><span class="sxs-lookup"><span data-stu-id="209f3-136">Compose</span></span><br><span data-ttu-id="209f3-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-137">Read</span></span> | <span data-ttu-id="209f3-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-138">String</span></span> | [<span data-ttu-id="209f3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="209f3-140">EventType</span><span class="sxs-lookup"><span data-stu-id="209f3-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="209f3-141">Composition</span><span class="sxs-lookup"><span data-stu-id="209f3-141">Compose</span></span><br><span data-ttu-id="209f3-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-142">Read</span></span> | <span data-ttu-id="209f3-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-143">String</span></span> | [<span data-ttu-id="209f3-144">1,5</span><span class="sxs-lookup"><span data-stu-id="209f3-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="209f3-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="209f3-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="209f3-146">Composition</span><span class="sxs-lookup"><span data-stu-id="209f3-146">Compose</span></span><br><span data-ttu-id="209f3-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-147">Read</span></span> | <span data-ttu-id="209f3-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-148">String</span></span> | [<span data-ttu-id="209f3-149">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="209f3-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="209f3-150">Namespaces</span></span>

<span data-ttu-id="209f3-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="209f3-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="209f3-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="209f3-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="209f3-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="209f3-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="209f3-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="209f3-155">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-155">Type</span></span>

*   <span data-ttu-id="209f3-156">String</span><span class="sxs-lookup"><span data-stu-id="209f3-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="209f3-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="209f3-157">Properties:</span></span>

|<span data-ttu-id="209f3-158">Nom</span><span class="sxs-lookup"><span data-stu-id="209f3-158">Name</span></span>| <span data-ttu-id="209f3-159">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-159">Type</span></span>| <span data-ttu-id="209f3-160">Description</span><span class="sxs-lookup"><span data-stu-id="209f3-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="209f3-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-161">String</span></span>|<span data-ttu-id="209f3-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="209f3-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="209f3-163">String</span><span class="sxs-lookup"><span data-stu-id="209f3-163">String</span></span>|<span data-ttu-id="209f3-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="209f3-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="209f3-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="209f3-165">Requirements</span></span>

|<span data-ttu-id="209f3-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="209f3-166">Requirement</span></span>| <span data-ttu-id="209f3-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="209f3-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="209f3-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="209f3-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="209f3-169">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-169">1.1</span></span>|
|[<span data-ttu-id="209f3-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="209f3-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="209f3-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="209f3-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-172">CoercionType: String</span></span>

<span data-ttu-id="209f3-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="209f3-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="209f3-174">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-174">Type</span></span>

*   <span data-ttu-id="209f3-175">String</span><span class="sxs-lookup"><span data-stu-id="209f3-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="209f3-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="209f3-176">Properties:</span></span>

|<span data-ttu-id="209f3-177">Nom</span><span class="sxs-lookup"><span data-stu-id="209f3-177">Name</span></span>| <span data-ttu-id="209f3-178">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-178">Type</span></span>| <span data-ttu-id="209f3-179">Description</span><span class="sxs-lookup"><span data-stu-id="209f3-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="209f3-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-180">String</span></span>|<span data-ttu-id="209f3-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="209f3-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="209f3-182">String</span><span class="sxs-lookup"><span data-stu-id="209f3-182">String</span></span>|<span data-ttu-id="209f3-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="209f3-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="209f3-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="209f3-184">Requirements</span></span>

|<span data-ttu-id="209f3-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="209f3-185">Requirement</span></span>| <span data-ttu-id="209f3-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="209f3-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="209f3-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="209f3-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="209f3-188">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-188">1.1</span></span>|
|[<span data-ttu-id="209f3-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="209f3-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="209f3-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="209f3-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-191">EventType: String</span></span>

<span data-ttu-id="209f3-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="209f3-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="209f3-193">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-193">Type</span></span>

*   <span data-ttu-id="209f3-194">String</span><span class="sxs-lookup"><span data-stu-id="209f3-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="209f3-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="209f3-195">Properties:</span></span>

| <span data-ttu-id="209f3-196">Nom</span><span class="sxs-lookup"><span data-stu-id="209f3-196">Name</span></span> | <span data-ttu-id="209f3-197">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-197">Type</span></span> | <span data-ttu-id="209f3-198">Description</span><span class="sxs-lookup"><span data-stu-id="209f3-198">Description</span></span> | <span data-ttu-id="209f3-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="209f3-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="209f3-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-200">String</span></span> | <span data-ttu-id="209f3-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="209f3-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="209f3-202">1.7</span><span class="sxs-lookup"><span data-stu-id="209f3-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="209f3-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-203">String</span></span> | <span data-ttu-id="209f3-204">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="209f3-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="209f3-205">1,5</span><span class="sxs-lookup"><span data-stu-id="209f3-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="209f3-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-206">String</span></span> | <span data-ttu-id="209f3-207">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="209f3-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="209f3-208">1.7</span><span class="sxs-lookup"><span data-stu-id="209f3-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="209f3-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-209">String</span></span> | <span data-ttu-id="209f3-210">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="209f3-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="209f3-211">1.7</span><span class="sxs-lookup"><span data-stu-id="209f3-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="209f3-212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="209f3-212">Requirements</span></span>

|<span data-ttu-id="209f3-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="209f3-213">Requirement</span></span>| <span data-ttu-id="209f3-214">Valeur</span><span class="sxs-lookup"><span data-stu-id="209f3-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="209f3-215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="209f3-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="209f3-216">1,5</span><span class="sxs-lookup"><span data-stu-id="209f3-216">1.5</span></span> |
|[<span data-ttu-id="209f3-217">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="209f3-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="209f3-218">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="209f3-219">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-219">SourceProperty: String</span></span>

<span data-ttu-id="209f3-220">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="209f3-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="209f3-221">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-221">Type</span></span>

*   <span data-ttu-id="209f3-222">String</span><span class="sxs-lookup"><span data-stu-id="209f3-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="209f3-223">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="209f3-223">Properties:</span></span>

|<span data-ttu-id="209f3-224">Nom</span><span class="sxs-lookup"><span data-stu-id="209f3-224">Name</span></span>| <span data-ttu-id="209f3-225">Type</span><span class="sxs-lookup"><span data-stu-id="209f3-225">Type</span></span>| <span data-ttu-id="209f3-226">Description</span><span class="sxs-lookup"><span data-stu-id="209f3-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="209f3-227">Chaîne</span><span class="sxs-lookup"><span data-stu-id="209f3-227">String</span></span>|<span data-ttu-id="209f3-228">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="209f3-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="209f3-229">String</span><span class="sxs-lookup"><span data-stu-id="209f3-229">String</span></span>|<span data-ttu-id="209f3-230">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="209f3-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="209f3-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="209f3-231">Requirements</span></span>

|<span data-ttu-id="209f3-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="209f3-232">Requirement</span></span>| <span data-ttu-id="209f3-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="209f3-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="209f3-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="209f3-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="209f3-235">1.1</span><span class="sxs-lookup"><span data-stu-id="209f3-235">1.1</span></span>|
|[<span data-ttu-id="209f3-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="209f3-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="209f3-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="209f3-237">Compose or Read</span></span>|
