---
title: Office de noms - ensemble de conditions requises 1.9
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.9.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 203b901c619e19a8e5b9255e36274e2f6e1d1658
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590945"
---
# <a name="office-mailbox-requirement-set-19"></a><span data-ttu-id="e16c2-103">Office (ensemble de conditions requises de boîte aux lettres 1.9)</span><span class="sxs-lookup"><span data-stu-id="e16c2-103">Office (Mailbox requirement set 1.9)</span></span>

<span data-ttu-id="e16c2-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e16c2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e16c2-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e16c2-106">Requirements</span></span>

|<span data-ttu-id="e16c2-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e16c2-107">Requirement</span></span>| <span data-ttu-id="e16c2-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="e16c2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e16c2-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e16c2-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e16c2-110">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-110">1.1</span></span>|
|[<span data-ttu-id="e16c2-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e16c2-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e16c2-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e16c2-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="e16c2-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e16c2-113">Properties</span></span>

| <span data-ttu-id="e16c2-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="e16c2-114">Property</span></span> | <span data-ttu-id="e16c2-115">Modes</span><span class="sxs-lookup"><span data-stu-id="e16c2-115">Modes</span></span> | <span data-ttu-id="e16c2-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="e16c2-116">Return type</span></span> | <span data-ttu-id="e16c2-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="e16c2-117">Minimum</span></span><br><span data-ttu-id="e16c2-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e16c2-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e16c2-119">context</span><span class="sxs-lookup"><span data-stu-id="e16c2-119">context</span></span>](office.context.md) | <span data-ttu-id="e16c2-120">Composition</span><span class="sxs-lookup"><span data-stu-id="e16c2-120">Compose</span></span><br><span data-ttu-id="e16c2-121">Lire</span><span class="sxs-lookup"><span data-stu-id="e16c2-121">Read</span></span> | [<span data-ttu-id="e16c2-122">Context</span><span class="sxs-lookup"><span data-stu-id="e16c2-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e16c2-123">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="e16c2-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="e16c2-124">Enumerations</span></span>

| <span data-ttu-id="e16c2-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="e16c2-125">Enumeration</span></span> | <span data-ttu-id="e16c2-126">Modes</span><span class="sxs-lookup"><span data-stu-id="e16c2-126">Modes</span></span> | <span data-ttu-id="e16c2-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="e16c2-127">Return type</span></span> | <span data-ttu-id="e16c2-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="e16c2-128">Minimum</span></span><br><span data-ttu-id="e16c2-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e16c2-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e16c2-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e16c2-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e16c2-131">Composition</span><span class="sxs-lookup"><span data-stu-id="e16c2-131">Compose</span></span><br><span data-ttu-id="e16c2-132">Lire</span><span class="sxs-lookup"><span data-stu-id="e16c2-132">Read</span></span> | <span data-ttu-id="e16c2-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-133">String</span></span> | [<span data-ttu-id="e16c2-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e16c2-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e16c2-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e16c2-136">Composition</span><span class="sxs-lookup"><span data-stu-id="e16c2-136">Compose</span></span><br><span data-ttu-id="e16c2-137">Lire</span><span class="sxs-lookup"><span data-stu-id="e16c2-137">Read</span></span> | <span data-ttu-id="e16c2-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-138">String</span></span> | [<span data-ttu-id="e16c2-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e16c2-140">EventType</span><span class="sxs-lookup"><span data-stu-id="e16c2-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e16c2-141">Composition</span><span class="sxs-lookup"><span data-stu-id="e16c2-141">Compose</span></span><br><span data-ttu-id="e16c2-142">Lire</span><span class="sxs-lookup"><span data-stu-id="e16c2-142">Read</span></span> | <span data-ttu-id="e16c2-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-143">String</span></span> | [<span data-ttu-id="e16c2-144">1.5</span><span class="sxs-lookup"><span data-stu-id="e16c2-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e16c2-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e16c2-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e16c2-146">Composition</span><span class="sxs-lookup"><span data-stu-id="e16c2-146">Compose</span></span><br><span data-ttu-id="e16c2-147">Lire</span><span class="sxs-lookup"><span data-stu-id="e16c2-147">Read</span></span> | <span data-ttu-id="e16c2-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-148">String</span></span> | [<span data-ttu-id="e16c2-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="e16c2-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="e16c2-150">Namespaces</span></span>

<span data-ttu-id="e16c2-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="e16c2-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.9&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e16c2-152">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="e16c2-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e16c2-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="e16c2-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="e16c2-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="e16c2-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e16c2-155">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-155">Type</span></span>

*   <span data-ttu-id="e16c2-156">String</span><span class="sxs-lookup"><span data-stu-id="e16c2-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e16c2-157">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e16c2-157">Properties</span></span>

|<span data-ttu-id="e16c2-158">Nom</span><span class="sxs-lookup"><span data-stu-id="e16c2-158">Name</span></span>| <span data-ttu-id="e16c2-159">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-159">Type</span></span>| <span data-ttu-id="e16c2-160">Description</span><span class="sxs-lookup"><span data-stu-id="e16c2-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e16c2-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-161">String</span></span>|<span data-ttu-id="e16c2-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="e16c2-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e16c2-163">String</span><span class="sxs-lookup"><span data-stu-id="e16c2-163">String</span></span>|<span data-ttu-id="e16c2-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="e16c2-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e16c2-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e16c2-165">Requirements</span></span>

|<span data-ttu-id="e16c2-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e16c2-166">Requirement</span></span>| <span data-ttu-id="e16c2-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="e16c2-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e16c2-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e16c2-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e16c2-169">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-169">1.1</span></span>|
|[<span data-ttu-id="e16c2-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e16c2-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e16c2-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e16c2-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e16c2-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="e16c2-172">CoercionType: String</span></span>

<span data-ttu-id="e16c2-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e16c2-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e16c2-174">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-174">Type</span></span>

*   <span data-ttu-id="e16c2-175">String</span><span class="sxs-lookup"><span data-stu-id="e16c2-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e16c2-176">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e16c2-176">Properties</span></span>

|<span data-ttu-id="e16c2-177">Nom</span><span class="sxs-lookup"><span data-stu-id="e16c2-177">Name</span></span>| <span data-ttu-id="e16c2-178">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-178">Type</span></span>| <span data-ttu-id="e16c2-179">Description</span><span class="sxs-lookup"><span data-stu-id="e16c2-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e16c2-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-180">String</span></span>|<span data-ttu-id="e16c2-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="e16c2-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e16c2-182">String</span><span class="sxs-lookup"><span data-stu-id="e16c2-182">String</span></span>|<span data-ttu-id="e16c2-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="e16c2-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e16c2-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e16c2-184">Requirements</span></span>

|<span data-ttu-id="e16c2-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e16c2-185">Requirement</span></span>| <span data-ttu-id="e16c2-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="e16c2-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="e16c2-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e16c2-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e16c2-188">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-188">1.1</span></span>|
|[<span data-ttu-id="e16c2-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e16c2-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e16c2-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e16c2-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="e16c2-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="e16c2-191">EventType: String</span></span>

<span data-ttu-id="e16c2-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="e16c2-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e16c2-193">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-193">Type</span></span>

*   <span data-ttu-id="e16c2-194">String</span><span class="sxs-lookup"><span data-stu-id="e16c2-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e16c2-195">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e16c2-195">Properties</span></span>

| <span data-ttu-id="e16c2-196">Nom</span><span class="sxs-lookup"><span data-stu-id="e16c2-196">Name</span></span> | <span data-ttu-id="e16c2-197">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-197">Type</span></span> | <span data-ttu-id="e16c2-198">Description</span><span class="sxs-lookup"><span data-stu-id="e16c2-198">Description</span></span> | <span data-ttu-id="e16c2-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="e16c2-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="e16c2-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-200">String</span></span> | <span data-ttu-id="e16c2-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="e16c2-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e16c2-202">1.7</span><span class="sxs-lookup"><span data-stu-id="e16c2-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="e16c2-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-203">String</span></span> | <span data-ttu-id="e16c2-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="e16c2-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="e16c2-205">1.8</span><span class="sxs-lookup"><span data-stu-id="e16c2-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="e16c2-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-206">String</span></span> | <span data-ttu-id="e16c2-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="e16c2-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="e16c2-208">1.8</span><span class="sxs-lookup"><span data-stu-id="e16c2-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="e16c2-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-209">String</span></span> | <span data-ttu-id="e16c2-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="e16c2-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e16c2-211">1,5</span><span class="sxs-lookup"><span data-stu-id="e16c2-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e16c2-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-212">String</span></span> | <span data-ttu-id="e16c2-213">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="e16c2-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e16c2-214">1.7</span><span class="sxs-lookup"><span data-stu-id="e16c2-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e16c2-215">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-215">String</span></span> | <span data-ttu-id="e16c2-216">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="e16c2-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e16c2-217">1.7</span><span class="sxs-lookup"><span data-stu-id="e16c2-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e16c2-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e16c2-218">Requirements</span></span>

|<span data-ttu-id="e16c2-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e16c2-219">Requirement</span></span>| <span data-ttu-id="e16c2-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="e16c2-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="e16c2-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e16c2-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e16c2-222">1,5</span><span class="sxs-lookup"><span data-stu-id="e16c2-222">1.5</span></span> |
|[<span data-ttu-id="e16c2-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e16c2-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e16c2-224">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e16c2-224">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e16c2-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="e16c2-225">SourceProperty: String</span></span>

<span data-ttu-id="e16c2-226">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e16c2-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e16c2-227">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-227">Type</span></span>

*   <span data-ttu-id="e16c2-228">String</span><span class="sxs-lookup"><span data-stu-id="e16c2-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e16c2-229">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e16c2-229">Properties</span></span>

|<span data-ttu-id="e16c2-230">Nom</span><span class="sxs-lookup"><span data-stu-id="e16c2-230">Name</span></span>| <span data-ttu-id="e16c2-231">Type</span><span class="sxs-lookup"><span data-stu-id="e16c2-231">Type</span></span>| <span data-ttu-id="e16c2-232">Description</span><span class="sxs-lookup"><span data-stu-id="e16c2-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e16c2-233">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e16c2-233">String</span></span>|<span data-ttu-id="e16c2-234">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="e16c2-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e16c2-235">String</span><span class="sxs-lookup"><span data-stu-id="e16c2-235">String</span></span>|<span data-ttu-id="e16c2-236">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="e16c2-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e16c2-237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e16c2-237">Requirements</span></span>

|<span data-ttu-id="e16c2-238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e16c2-238">Requirement</span></span>| <span data-ttu-id="e16c2-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="e16c2-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="e16c2-240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e16c2-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e16c2-241">1.1</span><span class="sxs-lookup"><span data-stu-id="e16c2-241">1.1</span></span>|
|[<span data-ttu-id="e16c2-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e16c2-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e16c2-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e16c2-243">Compose or Read</span></span>|
