---
title: Espace de noms Office – ensemble de conditions requises
description: ''
ms.date: 10/30/2019
localization_priority: Normal
ms.openlocfilehash: eae6f99d166695f24f4a94e89ea4b876bea080ef
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902101"
---
# <a name="office"></a><span data-ttu-id="a316e-102">Office</span><span class="sxs-lookup"><span data-stu-id="a316e-102">Office</span></span>

<span data-ttu-id="a316e-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="a316e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a316e-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a316e-105">Requirements</span></span>

|<span data-ttu-id="a316e-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a316e-106">Requirement</span></span>| <span data-ttu-id="a316e-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="a316e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="a316e-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a316e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a316e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="a316e-109">1.0</span></span>|
|[<span data-ttu-id="a316e-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a316e-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a316e-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a316e-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a316e-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="a316e-112">Members and methods</span></span>

| <span data-ttu-id="a316e-113">Membre</span><span class="sxs-lookup"><span data-stu-id="a316e-113">Member</span></span> | <span data-ttu-id="a316e-114">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a316e-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="a316e-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="a316e-116">Membre</span><span class="sxs-lookup"><span data-stu-id="a316e-116">Member</span></span> |
| [<span data-ttu-id="a316e-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="a316e-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="a316e-118">Membre</span><span class="sxs-lookup"><span data-stu-id="a316e-118">Member</span></span> |
| [<span data-ttu-id="a316e-119">EventType</span><span class="sxs-lookup"><span data-stu-id="a316e-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="a316e-120">Membre</span><span class="sxs-lookup"><span data-stu-id="a316e-120">Member</span></span> |
| [<span data-ttu-id="a316e-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="a316e-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="a316e-122">Membre</span><span class="sxs-lookup"><span data-stu-id="a316e-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="a316e-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="a316e-123">Namespaces</span></span>

<span data-ttu-id="a316e-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="a316e-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="a316e-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): inclut un certain nombre d’énumérations, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="a316e-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="a316e-126">Members</span><span class="sxs-lookup"><span data-stu-id="a316e-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="a316e-127">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="a316e-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="a316e-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="a316e-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="a316e-129">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-129">Type</span></span>

*   <span data-ttu-id="a316e-130">String</span><span class="sxs-lookup"><span data-stu-id="a316e-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a316e-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a316e-131">Properties:</span></span>

|<span data-ttu-id="a316e-132">Nom</span><span class="sxs-lookup"><span data-stu-id="a316e-132">Name</span></span>| <span data-ttu-id="a316e-133">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-133">Type</span></span>| <span data-ttu-id="a316e-134">Description</span><span class="sxs-lookup"><span data-stu-id="a316e-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="a316e-135">String</span><span class="sxs-lookup"><span data-stu-id="a316e-135">String</span></span>|<span data-ttu-id="a316e-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="a316e-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="a316e-137">String</span><span class="sxs-lookup"><span data-stu-id="a316e-137">String</span></span>|<span data-ttu-id="a316e-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="a316e-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a316e-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a316e-139">Requirements</span></span>

|<span data-ttu-id="a316e-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a316e-140">Requirement</span></span>| <span data-ttu-id="a316e-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="a316e-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="a316e-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a316e-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a316e-143">1.0</span><span class="sxs-lookup"><span data-stu-id="a316e-143">1.0</span></span>|
|[<span data-ttu-id="a316e-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a316e-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a316e-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a316e-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="a316e-146">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="a316e-146">CoercionType: String</span></span>

<span data-ttu-id="a316e-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a316e-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a316e-148">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-148">Type</span></span>

*   <span data-ttu-id="a316e-149">String</span><span class="sxs-lookup"><span data-stu-id="a316e-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a316e-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a316e-150">Properties:</span></span>

|<span data-ttu-id="a316e-151">Nom</span><span class="sxs-lookup"><span data-stu-id="a316e-151">Name</span></span>| <span data-ttu-id="a316e-152">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-152">Type</span></span>| <span data-ttu-id="a316e-153">Description</span><span class="sxs-lookup"><span data-stu-id="a316e-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="a316e-154">String</span><span class="sxs-lookup"><span data-stu-id="a316e-154">String</span></span>|<span data-ttu-id="a316e-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="a316e-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="a316e-156">String</span><span class="sxs-lookup"><span data-stu-id="a316e-156">String</span></span>|<span data-ttu-id="a316e-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="a316e-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a316e-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a316e-158">Requirements</span></span>

|<span data-ttu-id="a316e-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a316e-159">Requirement</span></span>| <span data-ttu-id="a316e-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="a316e-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="a316e-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a316e-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a316e-162">1.0</span><span class="sxs-lookup"><span data-stu-id="a316e-162">1.0</span></span>|
|[<span data-ttu-id="a316e-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a316e-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a316e-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a316e-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="a316e-165">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="a316e-165">EventType: String</span></span>

<span data-ttu-id="a316e-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="a316e-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="a316e-167">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-167">Type</span></span>

*   <span data-ttu-id="a316e-168">String</span><span class="sxs-lookup"><span data-stu-id="a316e-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a316e-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a316e-169">Properties:</span></span>

| <span data-ttu-id="a316e-170">Nom</span><span class="sxs-lookup"><span data-stu-id="a316e-170">Name</span></span> | <span data-ttu-id="a316e-171">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-171">Type</span></span> | <span data-ttu-id="a316e-172">Description</span><span class="sxs-lookup"><span data-stu-id="a316e-172">Description</span></span> | <span data-ttu-id="a316e-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="a316e-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="a316e-174">String</span><span class="sxs-lookup"><span data-stu-id="a316e-174">String</span></span> | <span data-ttu-id="a316e-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="a316e-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="a316e-176">1.7</span><span class="sxs-lookup"><span data-stu-id="a316e-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="a316e-177">String</span><span class="sxs-lookup"><span data-stu-id="a316e-177">String</span></span> | <span data-ttu-id="a316e-178">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="a316e-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="a316e-179">1.8</span><span class="sxs-lookup"><span data-stu-id="a316e-179">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="a316e-180">String</span><span class="sxs-lookup"><span data-stu-id="a316e-180">String</span></span> | <span data-ttu-id="a316e-181">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="a316e-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="a316e-182">1.8</span><span class="sxs-lookup"><span data-stu-id="a316e-182">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="a316e-183">String</span><span class="sxs-lookup"><span data-stu-id="a316e-183">String</span></span> | <span data-ttu-id="a316e-184">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="a316e-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="a316e-185">1,5</span><span class="sxs-lookup"><span data-stu-id="a316e-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="a316e-186">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a316e-186">String</span></span> | <span data-ttu-id="a316e-187">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="a316e-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="a316e-188">Aperçu</span><span class="sxs-lookup"><span data-stu-id="a316e-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="a316e-189">String</span><span class="sxs-lookup"><span data-stu-id="a316e-189">String</span></span> | <span data-ttu-id="a316e-190">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="a316e-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="a316e-191">1.7</span><span class="sxs-lookup"><span data-stu-id="a316e-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="a316e-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="a316e-192">String</span></span> | <span data-ttu-id="a316e-193">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="a316e-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="a316e-194">1.7</span><span class="sxs-lookup"><span data-stu-id="a316e-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="a316e-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a316e-195">Requirements</span></span>

|<span data-ttu-id="a316e-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a316e-196">Requirement</span></span>| <span data-ttu-id="a316e-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="a316e-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="a316e-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a316e-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a316e-199">1,5</span><span class="sxs-lookup"><span data-stu-id="a316e-199">1.5</span></span> |
|[<span data-ttu-id="a316e-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a316e-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a316e-201">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a316e-201">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="a316e-202">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="a316e-202">SourceProperty: String</span></span>

<span data-ttu-id="a316e-203">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="a316e-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="a316e-204">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-204">Type</span></span>

*   <span data-ttu-id="a316e-205">String</span><span class="sxs-lookup"><span data-stu-id="a316e-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="a316e-206">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="a316e-206">Properties:</span></span>

|<span data-ttu-id="a316e-207">Nom</span><span class="sxs-lookup"><span data-stu-id="a316e-207">Name</span></span>| <span data-ttu-id="a316e-208">Type</span><span class="sxs-lookup"><span data-stu-id="a316e-208">Type</span></span>| <span data-ttu-id="a316e-209">Description</span><span class="sxs-lookup"><span data-stu-id="a316e-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="a316e-210">String</span><span class="sxs-lookup"><span data-stu-id="a316e-210">String</span></span>|<span data-ttu-id="a316e-211">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="a316e-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="a316e-212">String</span><span class="sxs-lookup"><span data-stu-id="a316e-212">String</span></span>|<span data-ttu-id="a316e-213">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="a316e-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="a316e-214">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="a316e-214">Requirements</span></span>

|<span data-ttu-id="a316e-215">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="a316e-215">Requirement</span></span>| <span data-ttu-id="a316e-216">Valeur</span><span class="sxs-lookup"><span data-stu-id="a316e-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="a316e-217">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="a316e-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a316e-218">1.0</span><span class="sxs-lookup"><span data-stu-id="a316e-218">1.0</span></span>|
|[<span data-ttu-id="a316e-219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="a316e-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="a316e-220">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="a316e-220">Compose or Read</span></span>|
