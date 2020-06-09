---
title: Espace de noms Office-ensemble de conditions requises 1,7
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 718de46689fc2fcb52ad455763581ecab06a4c39
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612198"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="8affe-103">Office (boîte aux lettres requise définie sur 1,7)</span><span class="sxs-lookup"><span data-stu-id="8affe-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="8affe-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="8affe-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8affe-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8affe-106">Requirements</span></span>

|<span data-ttu-id="8affe-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8affe-107">Requirement</span></span>| <span data-ttu-id="8affe-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="8affe-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="8affe-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8affe-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8affe-110">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-110">1.1</span></span>|
|[<span data-ttu-id="8affe-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8affe-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8affe-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8affe-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8affe-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="8affe-113">Properties</span></span>

| <span data-ttu-id="8affe-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="8affe-114">Property</span></span> | <span data-ttu-id="8affe-115">Modes</span><span class="sxs-lookup"><span data-stu-id="8affe-115">Modes</span></span> | <span data-ttu-id="8affe-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="8affe-116">Return type</span></span> | <span data-ttu-id="8affe-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="8affe-117">Minimum</span></span><br><span data-ttu-id="8affe-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="8affe-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8affe-119">context</span><span class="sxs-lookup"><span data-stu-id="8affe-119">context</span></span>](office.context.md) | <span data-ttu-id="8affe-120">Composition</span><span class="sxs-lookup"><span data-stu-id="8affe-120">Compose</span></span><br><span data-ttu-id="8affe-121">Read</span><span class="sxs-lookup"><span data-stu-id="8affe-121">Read</span></span> | [<span data-ttu-id="8affe-122">Context</span><span class="sxs-lookup"><span data-stu-id="8affe-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="8affe-123">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="8affe-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="8affe-124">Enumerations</span></span>

| <span data-ttu-id="8affe-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="8affe-125">Enumeration</span></span> | <span data-ttu-id="8affe-126">Modes</span><span class="sxs-lookup"><span data-stu-id="8affe-126">Modes</span></span> | <span data-ttu-id="8affe-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="8affe-127">Return type</span></span> | <span data-ttu-id="8affe-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="8affe-128">Minimum</span></span><br><span data-ttu-id="8affe-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="8affe-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8affe-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8affe-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8affe-131">Composition</span><span class="sxs-lookup"><span data-stu-id="8affe-131">Compose</span></span><br><span data-ttu-id="8affe-132">Read</span><span class="sxs-lookup"><span data-stu-id="8affe-132">Read</span></span> | <span data-ttu-id="8affe-133">String</span><span class="sxs-lookup"><span data-stu-id="8affe-133">String</span></span> | [<span data-ttu-id="8affe-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8affe-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8affe-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8affe-136">Composition</span><span class="sxs-lookup"><span data-stu-id="8affe-136">Compose</span></span><br><span data-ttu-id="8affe-137">Read</span><span class="sxs-lookup"><span data-stu-id="8affe-137">Read</span></span> | <span data-ttu-id="8affe-138">String</span><span class="sxs-lookup"><span data-stu-id="8affe-138">String</span></span> | [<span data-ttu-id="8affe-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8affe-140">EventType</span><span class="sxs-lookup"><span data-stu-id="8affe-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8affe-141">Composition</span><span class="sxs-lookup"><span data-stu-id="8affe-141">Compose</span></span><br><span data-ttu-id="8affe-142">Read</span><span class="sxs-lookup"><span data-stu-id="8affe-142">Read</span></span> | <span data-ttu-id="8affe-143">String</span><span class="sxs-lookup"><span data-stu-id="8affe-143">String</span></span> | [<span data-ttu-id="8affe-144">1,5</span><span class="sxs-lookup"><span data-stu-id="8affe-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="8affe-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8affe-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8affe-146">Composition</span><span class="sxs-lookup"><span data-stu-id="8affe-146">Compose</span></span><br><span data-ttu-id="8affe-147">Read</span><span class="sxs-lookup"><span data-stu-id="8affe-147">Read</span></span> | <span data-ttu-id="8affe-148">String</span><span class="sxs-lookup"><span data-stu-id="8affe-148">String</span></span> | [<span data-ttu-id="8affe-149">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="8affe-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="8affe-150">Namespaces</span></span>

<span data-ttu-id="8affe-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="8affe-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="8affe-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="8affe-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="8affe-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="8affe-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="8affe-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8affe-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8affe-155">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-155">Type</span></span>

*   <span data-ttu-id="8affe-156">String</span><span class="sxs-lookup"><span data-stu-id="8affe-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8affe-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8affe-157">Properties:</span></span>

|<span data-ttu-id="8affe-158">Nom</span><span class="sxs-lookup"><span data-stu-id="8affe-158">Name</span></span>| <span data-ttu-id="8affe-159">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-159">Type</span></span>| <span data-ttu-id="8affe-160">Description</span><span class="sxs-lookup"><span data-stu-id="8affe-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8affe-161">String</span><span class="sxs-lookup"><span data-stu-id="8affe-161">String</span></span>|<span data-ttu-id="8affe-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="8affe-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8affe-163">String</span><span class="sxs-lookup"><span data-stu-id="8affe-163">String</span></span>|<span data-ttu-id="8affe-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="8affe-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8affe-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8affe-165">Requirements</span></span>

|<span data-ttu-id="8affe-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8affe-166">Requirement</span></span>| <span data-ttu-id="8affe-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="8affe-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="8affe-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8affe-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8affe-169">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-169">1.1</span></span>|
|[<span data-ttu-id="8affe-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8affe-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8affe-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8affe-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="8affe-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="8affe-172">CoercionType: String</span></span>

<span data-ttu-id="8affe-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8affe-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8affe-174">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-174">Type</span></span>

*   <span data-ttu-id="8affe-175">String</span><span class="sxs-lookup"><span data-stu-id="8affe-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8affe-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8affe-176">Properties:</span></span>

|<span data-ttu-id="8affe-177">Nom</span><span class="sxs-lookup"><span data-stu-id="8affe-177">Name</span></span>| <span data-ttu-id="8affe-178">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-178">Type</span></span>| <span data-ttu-id="8affe-179">Description</span><span class="sxs-lookup"><span data-stu-id="8affe-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8affe-180">String</span><span class="sxs-lookup"><span data-stu-id="8affe-180">String</span></span>|<span data-ttu-id="8affe-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="8affe-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8affe-182">String</span><span class="sxs-lookup"><span data-stu-id="8affe-182">String</span></span>|<span data-ttu-id="8affe-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="8affe-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8affe-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8affe-184">Requirements</span></span>

|<span data-ttu-id="8affe-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8affe-185">Requirement</span></span>| <span data-ttu-id="8affe-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="8affe-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="8affe-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8affe-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8affe-188">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-188">1.1</span></span>|
|[<span data-ttu-id="8affe-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8affe-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8affe-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8affe-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="8affe-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="8affe-191">EventType: String</span></span>

<span data-ttu-id="8affe-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="8affe-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8affe-193">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-193">Type</span></span>

*   <span data-ttu-id="8affe-194">String</span><span class="sxs-lookup"><span data-stu-id="8affe-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8affe-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8affe-195">Properties:</span></span>

| <span data-ttu-id="8affe-196">Nom</span><span class="sxs-lookup"><span data-stu-id="8affe-196">Name</span></span> | <span data-ttu-id="8affe-197">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-197">Type</span></span> | <span data-ttu-id="8affe-198">Description</span><span class="sxs-lookup"><span data-stu-id="8affe-198">Description</span></span> | <span data-ttu-id="8affe-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="8affe-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="8affe-200">String</span><span class="sxs-lookup"><span data-stu-id="8affe-200">String</span></span> | <span data-ttu-id="8affe-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="8affe-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="8affe-202">1.7</span><span class="sxs-lookup"><span data-stu-id="8affe-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="8affe-203">String</span><span class="sxs-lookup"><span data-stu-id="8affe-203">String</span></span> | <span data-ttu-id="8affe-204">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="8affe-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="8affe-205">1,5</span><span class="sxs-lookup"><span data-stu-id="8affe-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="8affe-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8affe-206">String</span></span> | <span data-ttu-id="8affe-207">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="8affe-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="8affe-208">1.7</span><span class="sxs-lookup"><span data-stu-id="8affe-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="8affe-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8affe-209">String</span></span> | <span data-ttu-id="8affe-210">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="8affe-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="8affe-211">1.7</span><span class="sxs-lookup"><span data-stu-id="8affe-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8affe-212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8affe-212">Requirements</span></span>

|<span data-ttu-id="8affe-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8affe-213">Requirement</span></span>| <span data-ttu-id="8affe-214">Valeur</span><span class="sxs-lookup"><span data-stu-id="8affe-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="8affe-215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8affe-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8affe-216">1,5</span><span class="sxs-lookup"><span data-stu-id="8affe-216">1.5</span></span> |
|[<span data-ttu-id="8affe-217">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8affe-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8affe-218">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8affe-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="8affe-219">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="8affe-219">SourceProperty: String</span></span>

<span data-ttu-id="8affe-220">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8affe-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8affe-221">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-221">Type</span></span>

*   <span data-ttu-id="8affe-222">String</span><span class="sxs-lookup"><span data-stu-id="8affe-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8affe-223">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8affe-223">Properties:</span></span>

|<span data-ttu-id="8affe-224">Nom</span><span class="sxs-lookup"><span data-stu-id="8affe-224">Name</span></span>| <span data-ttu-id="8affe-225">Type</span><span class="sxs-lookup"><span data-stu-id="8affe-225">Type</span></span>| <span data-ttu-id="8affe-226">Description</span><span class="sxs-lookup"><span data-stu-id="8affe-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8affe-227">String</span><span class="sxs-lookup"><span data-stu-id="8affe-227">String</span></span>|<span data-ttu-id="8affe-228">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="8affe-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8affe-229">String</span><span class="sxs-lookup"><span data-stu-id="8affe-229">String</span></span>|<span data-ttu-id="8affe-230">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="8affe-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8affe-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8affe-231">Requirements</span></span>

|<span data-ttu-id="8affe-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8affe-232">Requirement</span></span>| <span data-ttu-id="8affe-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="8affe-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="8affe-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8affe-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8affe-235">1.1</span><span class="sxs-lookup"><span data-stu-id="8affe-235">1.1</span></span>|
|[<span data-ttu-id="8affe-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8affe-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8affe-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8affe-237">Compose or Read</span></span>|
