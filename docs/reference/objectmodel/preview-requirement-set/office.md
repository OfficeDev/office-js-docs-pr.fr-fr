---
title: Espace de noms Office – ensemble de conditions requises
description: Les membres d’espace de noms Office disponibles pour les compléments Outlook à l’aide de l’ensemble de conditions requises d’API de boîte aux lettres.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: d72e5c78a7fd8d3c00b8f84e7d9b05ee6defc0c5
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890857"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="59ffb-103">Office (préversion de la boîte aux lettres requise)</span><span class="sxs-lookup"><span data-stu-id="59ffb-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="59ffb-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="59ffb-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="59ffb-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="59ffb-106">Requirements</span></span>

|<span data-ttu-id="59ffb-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="59ffb-107">Requirement</span></span>| <span data-ttu-id="59ffb-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="59ffb-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="59ffb-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="59ffb-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="59ffb-110">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-110">1.1</span></span>|
|[<span data-ttu-id="59ffb-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="59ffb-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="59ffb-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="59ffb-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="59ffb-113">Properties</span></span>

| <span data-ttu-id="59ffb-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="59ffb-114">Property</span></span> | <span data-ttu-id="59ffb-115">Modes</span><span class="sxs-lookup"><span data-stu-id="59ffb-115">Modes</span></span> | <span data-ttu-id="59ffb-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="59ffb-116">Return type</span></span> | <span data-ttu-id="59ffb-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="59ffb-117">Minimum</span></span><br><span data-ttu-id="59ffb-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="59ffb-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="59ffb-119">context</span><span class="sxs-lookup"><span data-stu-id="59ffb-119">context</span></span>](office.context.md) | <span data-ttu-id="59ffb-120">Composition</span><span class="sxs-lookup"><span data-stu-id="59ffb-120">Compose</span></span><br><span data-ttu-id="59ffb-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-121">Read</span></span> | [<span data-ttu-id="59ffb-122">Context</span><span class="sxs-lookup"><span data-stu-id="59ffb-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="59ffb-123">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="59ffb-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="59ffb-124">Enumerations</span></span>

| <span data-ttu-id="59ffb-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="59ffb-125">Enumeration</span></span> | <span data-ttu-id="59ffb-126">Modes</span><span class="sxs-lookup"><span data-stu-id="59ffb-126">Modes</span></span> | <span data-ttu-id="59ffb-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="59ffb-127">Return type</span></span> | <span data-ttu-id="59ffb-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="59ffb-128">Minimum</span></span><br><span data-ttu-id="59ffb-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="59ffb-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="59ffb-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="59ffb-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="59ffb-131">Composition</span><span class="sxs-lookup"><span data-stu-id="59ffb-131">Compose</span></span><br><span data-ttu-id="59ffb-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-132">Read</span></span> | <span data-ttu-id="59ffb-133">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-133">String</span></span> | [<span data-ttu-id="59ffb-134">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="59ffb-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="59ffb-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="59ffb-136">Composition</span><span class="sxs-lookup"><span data-stu-id="59ffb-136">Compose</span></span><br><span data-ttu-id="59ffb-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-137">Read</span></span> | <span data-ttu-id="59ffb-138">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-138">String</span></span> | [<span data-ttu-id="59ffb-139">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="59ffb-140">EventType</span><span class="sxs-lookup"><span data-stu-id="59ffb-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="59ffb-141">Composition</span><span class="sxs-lookup"><span data-stu-id="59ffb-141">Compose</span></span><br><span data-ttu-id="59ffb-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-142">Read</span></span> | <span data-ttu-id="59ffb-143">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-143">String</span></span> | [<span data-ttu-id="59ffb-144">1,5</span><span class="sxs-lookup"><span data-stu-id="59ffb-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="59ffb-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="59ffb-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="59ffb-146">Composition</span><span class="sxs-lookup"><span data-stu-id="59ffb-146">Compose</span></span><br><span data-ttu-id="59ffb-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-147">Read</span></span> | <span data-ttu-id="59ffb-148">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-148">String</span></span> | [<span data-ttu-id="59ffb-149">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="59ffb-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="59ffb-150">Namespaces</span></span>

<span data-ttu-id="59ffb-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="59ffb-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="59ffb-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="59ffb-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="59ffb-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="59ffb-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="59ffb-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="59ffb-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="59ffb-155">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-155">Type</span></span>

*   <span data-ttu-id="59ffb-156">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="59ffb-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="59ffb-157">Properties:</span></span>

|<span data-ttu-id="59ffb-158">Nom</span><span class="sxs-lookup"><span data-stu-id="59ffb-158">Name</span></span>| <span data-ttu-id="59ffb-159">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-159">Type</span></span>| <span data-ttu-id="59ffb-160">Description</span><span class="sxs-lookup"><span data-stu-id="59ffb-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="59ffb-161">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-161">String</span></span>|<span data-ttu-id="59ffb-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="59ffb-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="59ffb-163">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-163">String</span></span>|<span data-ttu-id="59ffb-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="59ffb-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59ffb-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="59ffb-165">Requirements</span></span>

|<span data-ttu-id="59ffb-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="59ffb-166">Requirement</span></span>| <span data-ttu-id="59ffb-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="59ffb-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="59ffb-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="59ffb-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="59ffb-169">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-169">1.1</span></span>|
|[<span data-ttu-id="59ffb-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="59ffb-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="59ffb-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="59ffb-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="59ffb-172">CoercionType: String</span></span>

<span data-ttu-id="59ffb-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="59ffb-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="59ffb-174">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-174">Type</span></span>

*   <span data-ttu-id="59ffb-175">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="59ffb-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="59ffb-176">Properties:</span></span>

|<span data-ttu-id="59ffb-177">Nom</span><span class="sxs-lookup"><span data-stu-id="59ffb-177">Name</span></span>| <span data-ttu-id="59ffb-178">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-178">Type</span></span>| <span data-ttu-id="59ffb-179">Description</span><span class="sxs-lookup"><span data-stu-id="59ffb-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="59ffb-180">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-180">String</span></span>|<span data-ttu-id="59ffb-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="59ffb-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="59ffb-182">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-182">String</span></span>|<span data-ttu-id="59ffb-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="59ffb-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59ffb-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="59ffb-184">Requirements</span></span>

|<span data-ttu-id="59ffb-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="59ffb-185">Requirement</span></span>| <span data-ttu-id="59ffb-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="59ffb-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="59ffb-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="59ffb-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="59ffb-188">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-188">1.1</span></span>|
|[<span data-ttu-id="59ffb-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="59ffb-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="59ffb-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="59ffb-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="59ffb-191">EventType: String</span></span>

<span data-ttu-id="59ffb-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="59ffb-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="59ffb-193">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-193">Type</span></span>

*   <span data-ttu-id="59ffb-194">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="59ffb-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="59ffb-195">Properties:</span></span>

| <span data-ttu-id="59ffb-196">Nom</span><span class="sxs-lookup"><span data-stu-id="59ffb-196">Name</span></span> | <span data-ttu-id="59ffb-197">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-197">Type</span></span> | <span data-ttu-id="59ffb-198">Description</span><span class="sxs-lookup"><span data-stu-id="59ffb-198">Description</span></span> | <span data-ttu-id="59ffb-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="59ffb-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="59ffb-200">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-200">String</span></span> | <span data-ttu-id="59ffb-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="59ffb-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="59ffb-202">1.7</span><span class="sxs-lookup"><span data-stu-id="59ffb-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="59ffb-203">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-203">String</span></span> | <span data-ttu-id="59ffb-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="59ffb-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="59ffb-205">1.8</span><span class="sxs-lookup"><span data-stu-id="59ffb-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="59ffb-206">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-206">String</span></span> | <span data-ttu-id="59ffb-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="59ffb-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="59ffb-208">1.8</span><span class="sxs-lookup"><span data-stu-id="59ffb-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="59ffb-209">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-209">String</span></span> | <span data-ttu-id="59ffb-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="59ffb-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="59ffb-211">1,5</span><span class="sxs-lookup"><span data-stu-id="59ffb-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="59ffb-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="59ffb-212">String</span></span> | <span data-ttu-id="59ffb-213">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="59ffb-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="59ffb-214">Aperçu</span><span class="sxs-lookup"><span data-stu-id="59ffb-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="59ffb-215">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-215">String</span></span> | <span data-ttu-id="59ffb-216">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="59ffb-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="59ffb-217">1.7</span><span class="sxs-lookup"><span data-stu-id="59ffb-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="59ffb-218">Chaîne</span><span class="sxs-lookup"><span data-stu-id="59ffb-218">String</span></span> | <span data-ttu-id="59ffb-219">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="59ffb-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="59ffb-220">1.7</span><span class="sxs-lookup"><span data-stu-id="59ffb-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="59ffb-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="59ffb-221">Requirements</span></span>

|<span data-ttu-id="59ffb-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="59ffb-222">Requirement</span></span>| <span data-ttu-id="59ffb-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="59ffb-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="59ffb-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="59ffb-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="59ffb-225">1,5</span><span class="sxs-lookup"><span data-stu-id="59ffb-225">1.5</span></span> |
|[<span data-ttu-id="59ffb-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="59ffb-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="59ffb-227">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="59ffb-228">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="59ffb-228">SourceProperty: String</span></span>

<span data-ttu-id="59ffb-229">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="59ffb-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="59ffb-230">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-230">Type</span></span>

*   <span data-ttu-id="59ffb-231">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="59ffb-232">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="59ffb-232">Properties:</span></span>

|<span data-ttu-id="59ffb-233">Nom</span><span class="sxs-lookup"><span data-stu-id="59ffb-233">Name</span></span>| <span data-ttu-id="59ffb-234">Type</span><span class="sxs-lookup"><span data-stu-id="59ffb-234">Type</span></span>| <span data-ttu-id="59ffb-235">Description</span><span class="sxs-lookup"><span data-stu-id="59ffb-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="59ffb-236">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-236">String</span></span>|<span data-ttu-id="59ffb-237">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="59ffb-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="59ffb-238">String</span><span class="sxs-lookup"><span data-stu-id="59ffb-238">String</span></span>|<span data-ttu-id="59ffb-239">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="59ffb-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="59ffb-240">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="59ffb-240">Requirements</span></span>

|<span data-ttu-id="59ffb-241">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="59ffb-241">Requirement</span></span>| <span data-ttu-id="59ffb-242">Valeur</span><span class="sxs-lookup"><span data-stu-id="59ffb-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="59ffb-243">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="59ffb-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="59ffb-244">1.1</span><span class="sxs-lookup"><span data-stu-id="59ffb-244">1.1</span></span>|
|[<span data-ttu-id="59ffb-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="59ffb-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="59ffb-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="59ffb-246">Compose or Read</span></span>|
