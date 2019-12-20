---
title: Espace de noms Office – ensemble de conditions requises
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ef9634058fcdc633e9ad3a0adb74c4abebf8038b
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40815059"
---
# <a name="office"></a><span data-ttu-id="2fc2e-102">Office</span><span class="sxs-lookup"><span data-stu-id="2fc2e-102">Office</span></span>

<span data-ttu-id="2fc2e-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="2fc2e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2fc2e-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fc2e-105">Requirements</span></span>

|<span data-ttu-id="2fc2e-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fc2e-106">Requirement</span></span>| <span data-ttu-id="2fc2e-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fc2e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fc2e-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fc2e-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2fc2e-109">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-109">1.1</span></span>|
|[<span data-ttu-id="2fc2e-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fc2e-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fc2e-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2fc2e-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="2fc2e-112">Properties</span></span>

| <span data-ttu-id="2fc2e-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="2fc2e-113">Property</span></span> | <span data-ttu-id="2fc2e-114">Modes</span><span class="sxs-lookup"><span data-stu-id="2fc2e-114">Modes</span></span> | <span data-ttu-id="2fc2e-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="2fc2e-115">Return type</span></span> | <span data-ttu-id="2fc2e-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="2fc2e-116">Minimum</span></span><br><span data-ttu-id="2fc2e-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fc2e-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2fc2e-118">context</span><span class="sxs-lookup"><span data-stu-id="2fc2e-118">context</span></span>](office.context.md) | <span data-ttu-id="2fc2e-119">Composition</span><span class="sxs-lookup"><span data-stu-id="2fc2e-119">Compose</span></span><br><span data-ttu-id="2fc2e-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-120">Read</span></span> | [<span data-ttu-id="2fc2e-121">Context</span><span class="sxs-lookup"><span data-stu-id="2fc2e-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview) | [<span data-ttu-id="2fc2e-122">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2fc2e-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="2fc2e-123">Enumerations</span></span>

| <span data-ttu-id="2fc2e-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="2fc2e-124">Enumeration</span></span> | <span data-ttu-id="2fc2e-125">Modes</span><span class="sxs-lookup"><span data-stu-id="2fc2e-125">Modes</span></span> | <span data-ttu-id="2fc2e-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="2fc2e-126">Return type</span></span> | <span data-ttu-id="2fc2e-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="2fc2e-127">Minimum</span></span><br><span data-ttu-id="2fc2e-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fc2e-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2fc2e-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2fc2e-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2fc2e-130">Composition</span><span class="sxs-lookup"><span data-stu-id="2fc2e-130">Compose</span></span><br><span data-ttu-id="2fc2e-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-131">Read</span></span> | <span data-ttu-id="2fc2e-132">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-132">String</span></span> | [<span data-ttu-id="2fc2e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2fc2e-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2fc2e-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2fc2e-135">Composition</span><span class="sxs-lookup"><span data-stu-id="2fc2e-135">Compose</span></span><br><span data-ttu-id="2fc2e-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-136">Read</span></span> | <span data-ttu-id="2fc2e-137">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-137">String</span></span> | [<span data-ttu-id="2fc2e-138">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2fc2e-139">EventType</span><span class="sxs-lookup"><span data-stu-id="2fc2e-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2fc2e-140">Composition</span><span class="sxs-lookup"><span data-stu-id="2fc2e-140">Compose</span></span><br><span data-ttu-id="2fc2e-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-141">Read</span></span> | <span data-ttu-id="2fc2e-142">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-142">String</span></span> | [<span data-ttu-id="2fc2e-143">1,5</span><span class="sxs-lookup"><span data-stu-id="2fc2e-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2fc2e-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2fc2e-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2fc2e-145">Composition</span><span class="sxs-lookup"><span data-stu-id="2fc2e-145">Compose</span></span><br><span data-ttu-id="2fc2e-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-146">Read</span></span> | <span data-ttu-id="2fc2e-147">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-147">String</span></span> | [<span data-ttu-id="2fc2e-148">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2fc2e-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="2fc2e-149">Namespaces</span></span>

<span data-ttu-id="2fc2e-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2fc2e-151">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="2fc2e-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2fc2e-152">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="2fc2e-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="2fc2e-153">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2fc2e-154">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-154">Type</span></span>

*   <span data-ttu-id="2fc2e-155">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2fc2e-156">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2fc2e-156">Properties:</span></span>

|<span data-ttu-id="2fc2e-157">Nom</span><span class="sxs-lookup"><span data-stu-id="2fc2e-157">Name</span></span>| <span data-ttu-id="2fc2e-158">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-158">Type</span></span>| <span data-ttu-id="2fc2e-159">Description</span><span class="sxs-lookup"><span data-stu-id="2fc2e-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2fc2e-160">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-160">String</span></span>|<span data-ttu-id="2fc2e-161">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2fc2e-162">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-162">String</span></span>|<span data-ttu-id="2fc2e-163">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fc2e-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fc2e-164">Requirements</span></span>

|<span data-ttu-id="2fc2e-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fc2e-165">Requirement</span></span>| <span data-ttu-id="2fc2e-166">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fc2e-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fc2e-167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fc2e-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2fc2e-168">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-168">1.1</span></span>|
|[<span data-ttu-id="2fc2e-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fc2e-169">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fc2e-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2fc2e-171">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="2fc2e-171">CoercionType: String</span></span>

<span data-ttu-id="2fc2e-172">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2fc2e-173">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-173">Type</span></span>

*   <span data-ttu-id="2fc2e-174">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2fc2e-175">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2fc2e-175">Properties:</span></span>

|<span data-ttu-id="2fc2e-176">Nom</span><span class="sxs-lookup"><span data-stu-id="2fc2e-176">Name</span></span>| <span data-ttu-id="2fc2e-177">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-177">Type</span></span>| <span data-ttu-id="2fc2e-178">Description</span><span class="sxs-lookup"><span data-stu-id="2fc2e-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2fc2e-179">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-179">String</span></span>|<span data-ttu-id="2fc2e-180">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2fc2e-181">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-181">String</span></span>|<span data-ttu-id="2fc2e-182">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fc2e-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fc2e-183">Requirements</span></span>

|<span data-ttu-id="2fc2e-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fc2e-184">Requirement</span></span>| <span data-ttu-id="2fc2e-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fc2e-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fc2e-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fc2e-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2fc2e-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-187">1.1</span></span>|
|[<span data-ttu-id="2fc2e-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fc2e-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fc2e-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2fc2e-190">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="2fc2e-190">EventType: String</span></span>

<span data-ttu-id="2fc2e-191">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2fc2e-192">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-192">Type</span></span>

*   <span data-ttu-id="2fc2e-193">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2fc2e-194">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2fc2e-194">Properties:</span></span>

| <span data-ttu-id="2fc2e-195">Nom</span><span class="sxs-lookup"><span data-stu-id="2fc2e-195">Name</span></span> | <span data-ttu-id="2fc2e-196">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-196">Type</span></span> | <span data-ttu-id="2fc2e-197">Description</span><span class="sxs-lookup"><span data-stu-id="2fc2e-197">Description</span></span> | <span data-ttu-id="2fc2e-198">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="2fc2e-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="2fc2e-199">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-199">String</span></span> | <span data-ttu-id="2fc2e-200">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="2fc2e-201">1.7</span><span class="sxs-lookup"><span data-stu-id="2fc2e-201">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="2fc2e-202">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-202">String</span></span> | <span data-ttu-id="2fc2e-203">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-203">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="2fc2e-204">1.8</span><span class="sxs-lookup"><span data-stu-id="2fc2e-204">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="2fc2e-205">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-205">String</span></span> | <span data-ttu-id="2fc2e-206">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-206">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="2fc2e-207">1.8</span><span class="sxs-lookup"><span data-stu-id="2fc2e-207">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="2fc2e-208">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-208">String</span></span> | <span data-ttu-id="2fc2e-209">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-209">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2fc2e-210">1,5</span><span class="sxs-lookup"><span data-stu-id="2fc2e-210">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="2fc2e-211">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fc2e-211">String</span></span> | <span data-ttu-id="2fc2e-212">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-212">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="2fc2e-213">Aperçu</span><span class="sxs-lookup"><span data-stu-id="2fc2e-213">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="2fc2e-214">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-214">String</span></span> | <span data-ttu-id="2fc2e-215">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-215">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="2fc2e-216">1.7</span><span class="sxs-lookup"><span data-stu-id="2fc2e-216">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="2fc2e-217">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2fc2e-217">String</span></span> | <span data-ttu-id="2fc2e-218">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-218">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="2fc2e-219">1.7</span><span class="sxs-lookup"><span data-stu-id="2fc2e-219">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2fc2e-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fc2e-220">Requirements</span></span>

|<span data-ttu-id="2fc2e-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fc2e-221">Requirement</span></span>| <span data-ttu-id="2fc2e-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fc2e-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fc2e-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fc2e-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2fc2e-224">1,5</span><span class="sxs-lookup"><span data-stu-id="2fc2e-224">1.5</span></span> |
|[<span data-ttu-id="2fc2e-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fc2e-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fc2e-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-226">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2fc2e-227">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="2fc2e-227">SourceProperty: String</span></span>

<span data-ttu-id="2fc2e-228">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-228">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2fc2e-229">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-229">Type</span></span>

*   <span data-ttu-id="2fc2e-230">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-230">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2fc2e-231">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2fc2e-231">Properties:</span></span>

|<span data-ttu-id="2fc2e-232">Nom</span><span class="sxs-lookup"><span data-stu-id="2fc2e-232">Name</span></span>| <span data-ttu-id="2fc2e-233">Type</span><span class="sxs-lookup"><span data-stu-id="2fc2e-233">Type</span></span>| <span data-ttu-id="2fc2e-234">Description</span><span class="sxs-lookup"><span data-stu-id="2fc2e-234">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2fc2e-235">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-235">String</span></span>|<span data-ttu-id="2fc2e-236">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-236">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2fc2e-237">String</span><span class="sxs-lookup"><span data-stu-id="2fc2e-237">String</span></span>|<span data-ttu-id="2fc2e-238">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="2fc2e-238">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2fc2e-239">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2fc2e-239">Requirements</span></span>

|<span data-ttu-id="2fc2e-240">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2fc2e-240">Requirement</span></span>| <span data-ttu-id="2fc2e-241">Valeur</span><span class="sxs-lookup"><span data-stu-id="2fc2e-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="2fc2e-242">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2fc2e-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2fc2e-243">1.1</span><span class="sxs-lookup"><span data-stu-id="2fc2e-243">1.1</span></span>|
|[<span data-ttu-id="2fc2e-244">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2fc2e-244">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2fc2e-245">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2fc2e-245">Compose or Read</span></span>|
