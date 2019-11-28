---
title: Espace de noms Office – ensemble de conditions requises
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: bd37b1be4d77d73cb56b0b2593ccc57dea6cab27
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629229"
---
# <a name="office"></a><span data-ttu-id="05445-102">Office</span><span class="sxs-lookup"><span data-stu-id="05445-102">Office</span></span>

<span data-ttu-id="05445-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="05445-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="05445-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="05445-105">Requirements</span></span>

|<span data-ttu-id="05445-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="05445-106">Requirement</span></span>| <span data-ttu-id="05445-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="05445-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="05445-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="05445-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="05445-109">1.0</span><span class="sxs-lookup"><span data-stu-id="05445-109">1.0</span></span>|
|[<span data-ttu-id="05445-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="05445-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05445-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="05445-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="05445-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="05445-112">Properties</span></span>

| <span data-ttu-id="05445-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="05445-113">Property</span></span> | <span data-ttu-id="05445-114">Modes</span><span class="sxs-lookup"><span data-stu-id="05445-114">Modes</span></span> | <span data-ttu-id="05445-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="05445-115">Return type</span></span> | <span data-ttu-id="05445-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="05445-116">Minimum</span></span><br><span data-ttu-id="05445-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="05445-117">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="05445-118">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="05445-118">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="05445-119">Composition</span><span class="sxs-lookup"><span data-stu-id="05445-119">Compose</span></span><br><span data-ttu-id="05445-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="05445-120">Read</span></span> | <span data-ttu-id="05445-121">String</span><span class="sxs-lookup"><span data-stu-id="05445-121">String</span></span> | <span data-ttu-id="05445-122">1.0</span><span class="sxs-lookup"><span data-stu-id="05445-122">1.0</span></span> |
| [<span data-ttu-id="05445-123">CoercionType</span><span class="sxs-lookup"><span data-stu-id="05445-123">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="05445-124">Composition</span><span class="sxs-lookup"><span data-stu-id="05445-124">Compose</span></span><br><span data-ttu-id="05445-125">Lecture</span><span class="sxs-lookup"><span data-stu-id="05445-125">Read</span></span> | <span data-ttu-id="05445-126">String</span><span class="sxs-lookup"><span data-stu-id="05445-126">String</span></span> | <span data-ttu-id="05445-127">1.0</span><span class="sxs-lookup"><span data-stu-id="05445-127">1.0</span></span> |
| [<span data-ttu-id="05445-128">EventType</span><span class="sxs-lookup"><span data-stu-id="05445-128">EventType</span></span>](#eventtype-string) | <span data-ttu-id="05445-129">Composition</span><span class="sxs-lookup"><span data-stu-id="05445-129">Compose</span></span><br><span data-ttu-id="05445-130">Lecture</span><span class="sxs-lookup"><span data-stu-id="05445-130">Read</span></span> | <span data-ttu-id="05445-131">String</span><span class="sxs-lookup"><span data-stu-id="05445-131">String</span></span> | <span data-ttu-id="05445-132">1,5</span><span class="sxs-lookup"><span data-stu-id="05445-132">1.5</span></span> |
| [<span data-ttu-id="05445-133">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="05445-133">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="05445-134">Composition</span><span class="sxs-lookup"><span data-stu-id="05445-134">Compose</span></span><br><span data-ttu-id="05445-135">Lecture</span><span class="sxs-lookup"><span data-stu-id="05445-135">Read</span></span> | <span data-ttu-id="05445-136">String</span><span class="sxs-lookup"><span data-stu-id="05445-136">String</span></span> | <span data-ttu-id="05445-137">1.0</span><span class="sxs-lookup"><span data-stu-id="05445-137">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="05445-138">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="05445-138">Namespaces</span></span>

<span data-ttu-id="05445-139">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="05445-139">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="05445-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): inclut un certain nombre d’énumérations, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="05445-140">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="property-details"></a><span data-ttu-id="05445-141">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="05445-141">Property details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="05445-142">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-142">AsyncResultStatus: String</span></span>

<span data-ttu-id="05445-143">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="05445-143">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="05445-144">Type</span><span class="sxs-lookup"><span data-stu-id="05445-144">Type</span></span>

*   <span data-ttu-id="05445-145">String</span><span class="sxs-lookup"><span data-stu-id="05445-145">String</span></span>

##### <a name="properties"></a><span data-ttu-id="05445-146">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="05445-146">Properties:</span></span>

|<span data-ttu-id="05445-147">Nom</span><span class="sxs-lookup"><span data-stu-id="05445-147">Name</span></span>| <span data-ttu-id="05445-148">Type</span><span class="sxs-lookup"><span data-stu-id="05445-148">Type</span></span>| <span data-ttu-id="05445-149">Description</span><span class="sxs-lookup"><span data-stu-id="05445-149">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="05445-150">Chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-150">String</span></span>|<span data-ttu-id="05445-151">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="05445-151">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="05445-152">String</span><span class="sxs-lookup"><span data-stu-id="05445-152">String</span></span>|<span data-ttu-id="05445-153">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="05445-153">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="05445-154">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="05445-154">Requirements</span></span>

|<span data-ttu-id="05445-155">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="05445-155">Requirement</span></span>| <span data-ttu-id="05445-156">Valeur</span><span class="sxs-lookup"><span data-stu-id="05445-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="05445-157">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="05445-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="05445-158">1.0</span><span class="sxs-lookup"><span data-stu-id="05445-158">1.0</span></span>|
|[<span data-ttu-id="05445-159">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="05445-159">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05445-160">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="05445-160">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="05445-161">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-161">CoercionType: String</span></span>

<span data-ttu-id="05445-162">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="05445-162">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="05445-163">Type</span><span class="sxs-lookup"><span data-stu-id="05445-163">Type</span></span>

*   <span data-ttu-id="05445-164">String</span><span class="sxs-lookup"><span data-stu-id="05445-164">String</span></span>

##### <a name="properties"></a><span data-ttu-id="05445-165">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="05445-165">Properties:</span></span>

|<span data-ttu-id="05445-166">Nom</span><span class="sxs-lookup"><span data-stu-id="05445-166">Name</span></span>| <span data-ttu-id="05445-167">Type</span><span class="sxs-lookup"><span data-stu-id="05445-167">Type</span></span>| <span data-ttu-id="05445-168">Description</span><span class="sxs-lookup"><span data-stu-id="05445-168">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="05445-169">Chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-169">String</span></span>|<span data-ttu-id="05445-170">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="05445-170">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="05445-171">String</span><span class="sxs-lookup"><span data-stu-id="05445-171">String</span></span>|<span data-ttu-id="05445-172">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="05445-172">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="05445-173">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="05445-173">Requirements</span></span>

|<span data-ttu-id="05445-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="05445-174">Requirement</span></span>| <span data-ttu-id="05445-175">Valeur</span><span class="sxs-lookup"><span data-stu-id="05445-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="05445-176">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="05445-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="05445-177">1.0</span><span class="sxs-lookup"><span data-stu-id="05445-177">1.0</span></span>|
|[<span data-ttu-id="05445-178">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="05445-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05445-179">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="05445-179">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="05445-180">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-180">EventType: String</span></span>

<span data-ttu-id="05445-181">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="05445-181">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="05445-182">Type</span><span class="sxs-lookup"><span data-stu-id="05445-182">Type</span></span>

*   <span data-ttu-id="05445-183">String</span><span class="sxs-lookup"><span data-stu-id="05445-183">String</span></span>

##### <a name="properties"></a><span data-ttu-id="05445-184">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="05445-184">Properties:</span></span>

| <span data-ttu-id="05445-185">Nom</span><span class="sxs-lookup"><span data-stu-id="05445-185">Name</span></span> | <span data-ttu-id="05445-186">Type</span><span class="sxs-lookup"><span data-stu-id="05445-186">Type</span></span> | <span data-ttu-id="05445-187">Description</span><span class="sxs-lookup"><span data-stu-id="05445-187">Description</span></span> | <span data-ttu-id="05445-188">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="05445-188">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="05445-189">String</span><span class="sxs-lookup"><span data-stu-id="05445-189">String</span></span> | <span data-ttu-id="05445-190">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="05445-190">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="05445-191">1.7</span><span class="sxs-lookup"><span data-stu-id="05445-191">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="05445-192">String</span><span class="sxs-lookup"><span data-stu-id="05445-192">String</span></span> | <span data-ttu-id="05445-193">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="05445-193">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="05445-194">1.8</span><span class="sxs-lookup"><span data-stu-id="05445-194">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="05445-195">String</span><span class="sxs-lookup"><span data-stu-id="05445-195">String</span></span> | <span data-ttu-id="05445-196">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="05445-196">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="05445-197">1.8</span><span class="sxs-lookup"><span data-stu-id="05445-197">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="05445-198">String</span><span class="sxs-lookup"><span data-stu-id="05445-198">String</span></span> | <span data-ttu-id="05445-199">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="05445-199">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="05445-200">1,5</span><span class="sxs-lookup"><span data-stu-id="05445-200">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="05445-201">Chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-201">String</span></span> | <span data-ttu-id="05445-202">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="05445-202">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="05445-203">Aperçu</span><span class="sxs-lookup"><span data-stu-id="05445-203">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="05445-204">String</span><span class="sxs-lookup"><span data-stu-id="05445-204">String</span></span> | <span data-ttu-id="05445-205">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="05445-205">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="05445-206">1.7</span><span class="sxs-lookup"><span data-stu-id="05445-206">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="05445-207">Chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-207">String</span></span> | <span data-ttu-id="05445-208">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="05445-208">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="05445-209">1.7</span><span class="sxs-lookup"><span data-stu-id="05445-209">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="05445-210">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="05445-210">Requirements</span></span>

|<span data-ttu-id="05445-211">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="05445-211">Requirement</span></span>| <span data-ttu-id="05445-212">Valeur</span><span class="sxs-lookup"><span data-stu-id="05445-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="05445-213">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="05445-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="05445-214">1,5</span><span class="sxs-lookup"><span data-stu-id="05445-214">1.5</span></span> |
|[<span data-ttu-id="05445-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="05445-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05445-216">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="05445-216">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="05445-217">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-217">SourceProperty: String</span></span>

<span data-ttu-id="05445-218">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="05445-218">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="05445-219">Type</span><span class="sxs-lookup"><span data-stu-id="05445-219">Type</span></span>

*   <span data-ttu-id="05445-220">String</span><span class="sxs-lookup"><span data-stu-id="05445-220">String</span></span>

##### <a name="properties"></a><span data-ttu-id="05445-221">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="05445-221">Properties:</span></span>

|<span data-ttu-id="05445-222">Nom</span><span class="sxs-lookup"><span data-stu-id="05445-222">Name</span></span>| <span data-ttu-id="05445-223">Type</span><span class="sxs-lookup"><span data-stu-id="05445-223">Type</span></span>| <span data-ttu-id="05445-224">Description</span><span class="sxs-lookup"><span data-stu-id="05445-224">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="05445-225">Chaîne</span><span class="sxs-lookup"><span data-stu-id="05445-225">String</span></span>|<span data-ttu-id="05445-226">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="05445-226">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="05445-227">String</span><span class="sxs-lookup"><span data-stu-id="05445-227">String</span></span>|<span data-ttu-id="05445-228">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="05445-228">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="05445-229">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="05445-229">Requirements</span></span>

|<span data-ttu-id="05445-230">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="05445-230">Requirement</span></span>| <span data-ttu-id="05445-231">Valeur</span><span class="sxs-lookup"><span data-stu-id="05445-231">Value</span></span>|
|---|---|
|[<span data-ttu-id="05445-232">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="05445-232">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="05445-233">1.0</span><span class="sxs-lookup"><span data-stu-id="05445-233">1.0</span></span>|
|[<span data-ttu-id="05445-234">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="05445-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="05445-235">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="05445-235">Compose or Read</span></span>|
