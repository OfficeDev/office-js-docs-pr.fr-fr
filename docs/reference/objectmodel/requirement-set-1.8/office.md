---
title: Espace de noms Office-ensemble de conditions requises 1,8
description: ''
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 91a0bef2a8280a068763c98b17644bd9268e2fb4
ms.sourcegitcommit: e989096f3d19761bf8477c585cde20b3f8e0b90d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/31/2019
ms.locfileid: "37902162"
---
# <a name="office"></a><span data-ttu-id="344c7-102">Office</span><span class="sxs-lookup"><span data-stu-id="344c7-102">Office</span></span>

<span data-ttu-id="344c7-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="344c7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="344c7-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="344c7-105">Requirements</span></span>

|<span data-ttu-id="344c7-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="344c7-106">Requirement</span></span>| <span data-ttu-id="344c7-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="344c7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="344c7-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="344c7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="344c7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="344c7-109">1.0</span></span>|
|[<span data-ttu-id="344c7-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="344c7-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="344c7-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="344c7-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="344c7-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="344c7-112">Members and methods</span></span>

| <span data-ttu-id="344c7-113">Membre</span><span class="sxs-lookup"><span data-stu-id="344c7-113">Member</span></span> | <span data-ttu-id="344c7-114">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="344c7-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="344c7-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="344c7-116">Membre</span><span class="sxs-lookup"><span data-stu-id="344c7-116">Member</span></span> |
| [<span data-ttu-id="344c7-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="344c7-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="344c7-118">Membre</span><span class="sxs-lookup"><span data-stu-id="344c7-118">Member</span></span> |
| [<span data-ttu-id="344c7-119">EventType</span><span class="sxs-lookup"><span data-stu-id="344c7-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="344c7-120">Membre</span><span class="sxs-lookup"><span data-stu-id="344c7-120">Member</span></span> |
| [<span data-ttu-id="344c7-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="344c7-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="344c7-122">Membre</span><span class="sxs-lookup"><span data-stu-id="344c7-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="344c7-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="344c7-123">Namespaces</span></span>

<span data-ttu-id="344c7-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="344c7-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="344c7-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclut un certain nombre d’énumérations, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="344c7-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="344c7-126">Members</span><span class="sxs-lookup"><span data-stu-id="344c7-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="344c7-127">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="344c7-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="344c7-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="344c7-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="344c7-129">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-129">Type</span></span>

*   <span data-ttu-id="344c7-130">String</span><span class="sxs-lookup"><span data-stu-id="344c7-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="344c7-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="344c7-131">Properties:</span></span>

|<span data-ttu-id="344c7-132">Nom</span><span class="sxs-lookup"><span data-stu-id="344c7-132">Name</span></span>| <span data-ttu-id="344c7-133">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-133">Type</span></span>| <span data-ttu-id="344c7-134">Description</span><span class="sxs-lookup"><span data-stu-id="344c7-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="344c7-135">String</span><span class="sxs-lookup"><span data-stu-id="344c7-135">String</span></span>|<span data-ttu-id="344c7-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="344c7-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="344c7-137">String</span><span class="sxs-lookup"><span data-stu-id="344c7-137">String</span></span>|<span data-ttu-id="344c7-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="344c7-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="344c7-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="344c7-139">Requirements</span></span>

|<span data-ttu-id="344c7-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="344c7-140">Requirement</span></span>| <span data-ttu-id="344c7-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="344c7-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="344c7-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="344c7-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="344c7-143">1.0</span><span class="sxs-lookup"><span data-stu-id="344c7-143">1.0</span></span>|
|[<span data-ttu-id="344c7-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="344c7-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="344c7-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="344c7-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="344c7-146">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="344c7-146">CoercionType: String</span></span>

<span data-ttu-id="344c7-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="344c7-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="344c7-148">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-148">Type</span></span>

*   <span data-ttu-id="344c7-149">String</span><span class="sxs-lookup"><span data-stu-id="344c7-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="344c7-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="344c7-150">Properties:</span></span>

|<span data-ttu-id="344c7-151">Nom</span><span class="sxs-lookup"><span data-stu-id="344c7-151">Name</span></span>| <span data-ttu-id="344c7-152">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-152">Type</span></span>| <span data-ttu-id="344c7-153">Description</span><span class="sxs-lookup"><span data-stu-id="344c7-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="344c7-154">String</span><span class="sxs-lookup"><span data-stu-id="344c7-154">String</span></span>|<span data-ttu-id="344c7-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="344c7-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="344c7-156">String</span><span class="sxs-lookup"><span data-stu-id="344c7-156">String</span></span>|<span data-ttu-id="344c7-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="344c7-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="344c7-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="344c7-158">Requirements</span></span>

|<span data-ttu-id="344c7-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="344c7-159">Requirement</span></span>| <span data-ttu-id="344c7-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="344c7-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="344c7-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="344c7-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="344c7-162">1.0</span><span class="sxs-lookup"><span data-stu-id="344c7-162">1.0</span></span>|
|[<span data-ttu-id="344c7-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="344c7-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="344c7-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="344c7-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="344c7-165">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="344c7-165">EventType: String</span></span>

<span data-ttu-id="344c7-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="344c7-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="344c7-167">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-167">Type</span></span>

*   <span data-ttu-id="344c7-168">String</span><span class="sxs-lookup"><span data-stu-id="344c7-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="344c7-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="344c7-169">Properties:</span></span>

| <span data-ttu-id="344c7-170">Nom</span><span class="sxs-lookup"><span data-stu-id="344c7-170">Name</span></span> | <span data-ttu-id="344c7-171">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-171">Type</span></span> | <span data-ttu-id="344c7-172">Description</span><span class="sxs-lookup"><span data-stu-id="344c7-172">Description</span></span> | <span data-ttu-id="344c7-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="344c7-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="344c7-174">String</span><span class="sxs-lookup"><span data-stu-id="344c7-174">String</span></span> | <span data-ttu-id="344c7-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="344c7-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="344c7-176">1.7</span><span class="sxs-lookup"><span data-stu-id="344c7-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="344c7-177">String</span><span class="sxs-lookup"><span data-stu-id="344c7-177">String</span></span> | <span data-ttu-id="344c7-178">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="344c7-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="344c7-179">1.8</span><span class="sxs-lookup"><span data-stu-id="344c7-179">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="344c7-180">String</span><span class="sxs-lookup"><span data-stu-id="344c7-180">String</span></span> | <span data-ttu-id="344c7-181">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="344c7-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="344c7-182">1.8</span><span class="sxs-lookup"><span data-stu-id="344c7-182">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="344c7-183">String</span><span class="sxs-lookup"><span data-stu-id="344c7-183">String</span></span> | <span data-ttu-id="344c7-184">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="344c7-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="344c7-185">1,5</span><span class="sxs-lookup"><span data-stu-id="344c7-185">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="344c7-186">Chaîne</span><span class="sxs-lookup"><span data-stu-id="344c7-186">String</span></span> | <span data-ttu-id="344c7-187">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="344c7-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="344c7-188">1.7</span><span class="sxs-lookup"><span data-stu-id="344c7-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="344c7-189">Chaîne</span><span class="sxs-lookup"><span data-stu-id="344c7-189">String</span></span> | <span data-ttu-id="344c7-190">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="344c7-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="344c7-191">1.7</span><span class="sxs-lookup"><span data-stu-id="344c7-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="344c7-192">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="344c7-192">Requirements</span></span>

|<span data-ttu-id="344c7-193">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="344c7-193">Requirement</span></span>| <span data-ttu-id="344c7-194">Valeur</span><span class="sxs-lookup"><span data-stu-id="344c7-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="344c7-195">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="344c7-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="344c7-196">1,5</span><span class="sxs-lookup"><span data-stu-id="344c7-196">1.5</span></span> |
|[<span data-ttu-id="344c7-197">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="344c7-197">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="344c7-198">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="344c7-198">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="344c7-199">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="344c7-199">SourceProperty: String</span></span>

<span data-ttu-id="344c7-200">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="344c7-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="344c7-201">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-201">Type</span></span>

*   <span data-ttu-id="344c7-202">String</span><span class="sxs-lookup"><span data-stu-id="344c7-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="344c7-203">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="344c7-203">Properties:</span></span>

|<span data-ttu-id="344c7-204">Nom</span><span class="sxs-lookup"><span data-stu-id="344c7-204">Name</span></span>| <span data-ttu-id="344c7-205">Type</span><span class="sxs-lookup"><span data-stu-id="344c7-205">Type</span></span>| <span data-ttu-id="344c7-206">Description</span><span class="sxs-lookup"><span data-stu-id="344c7-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="344c7-207">String</span><span class="sxs-lookup"><span data-stu-id="344c7-207">String</span></span>|<span data-ttu-id="344c7-208">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="344c7-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="344c7-209">String</span><span class="sxs-lookup"><span data-stu-id="344c7-209">String</span></span>|<span data-ttu-id="344c7-210">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="344c7-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="344c7-211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="344c7-211">Requirements</span></span>

|<span data-ttu-id="344c7-212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="344c7-212">Requirement</span></span>| <span data-ttu-id="344c7-213">Valeur</span><span class="sxs-lookup"><span data-stu-id="344c7-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="344c7-214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="344c7-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="344c7-215">1.0</span><span class="sxs-lookup"><span data-stu-id="344c7-215">1.0</span></span>|
|[<span data-ttu-id="344c7-216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="344c7-216">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="344c7-217">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="344c7-217">Compose or Read</span></span>|
