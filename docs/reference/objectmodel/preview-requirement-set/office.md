---
title: Espace de noms Office – ensemble de conditions requises
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: eb8ff0a755c1908d7b96438f96386056cc16b24f
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696434"
---
# <a name="office"></a><span data-ttu-id="4fe26-102">Office</span><span class="sxs-lookup"><span data-stu-id="4fe26-102">Office</span></span>

<span data-ttu-id="4fe26-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4fe26-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4fe26-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4fe26-105">Requirements</span></span>

|<span data-ttu-id="4fe26-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4fe26-106">Requirement</span></span>| <span data-ttu-id="4fe26-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4fe26-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fe26-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4fe26-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fe26-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4fe26-109">1.0</span></span>|
|[<span data-ttu-id="4fe26-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4fe26-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4fe26-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4fe26-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4fe26-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4fe26-112">Members and methods</span></span>

| <span data-ttu-id="4fe26-113">Membre</span><span class="sxs-lookup"><span data-stu-id="4fe26-113">Member</span></span> | <span data-ttu-id="4fe26-114">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4fe26-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4fe26-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4fe26-116">Member</span><span class="sxs-lookup"><span data-stu-id="4fe26-116">Member</span></span> |
| [<span data-ttu-id="4fe26-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4fe26-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4fe26-118">Member</span><span class="sxs-lookup"><span data-stu-id="4fe26-118">Member</span></span> |
| [<span data-ttu-id="4fe26-119">EventType</span><span class="sxs-lookup"><span data-stu-id="4fe26-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4fe26-120">Member</span><span class="sxs-lookup"><span data-stu-id="4fe26-120">Member</span></span> |
| [<span data-ttu-id="4fe26-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4fe26-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4fe26-122">Membre</span><span class="sxs-lookup"><span data-stu-id="4fe26-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4fe26-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4fe26-123">Namespaces</span></span>

<span data-ttu-id="4fe26-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4fe26-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4fe26-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): inclut un certain nombre d’énumérations, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="4fe26-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="4fe26-126">Members</span><span class="sxs-lookup"><span data-stu-id="4fe26-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="4fe26-127">AsyncResultStatus: chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="4fe26-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4fe26-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4fe26-129">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-129">Type</span></span>

*   <span data-ttu-id="4fe26-130">String</span><span class="sxs-lookup"><span data-stu-id="4fe26-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fe26-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4fe26-131">Properties:</span></span>

|<span data-ttu-id="4fe26-132">Nom</span><span class="sxs-lookup"><span data-stu-id="4fe26-132">Name</span></span>| <span data-ttu-id="4fe26-133">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-133">Type</span></span>| <span data-ttu-id="4fe26-134">Description</span><span class="sxs-lookup"><span data-stu-id="4fe26-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4fe26-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-135">String</span></span>|<span data-ttu-id="4fe26-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="4fe26-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4fe26-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-137">String</span></span>|<span data-ttu-id="4fe26-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="4fe26-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fe26-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4fe26-139">Requirements</span></span>

|<span data-ttu-id="4fe26-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4fe26-140">Requirement</span></span>| <span data-ttu-id="4fe26-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="4fe26-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fe26-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4fe26-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fe26-143">1.0</span><span class="sxs-lookup"><span data-stu-id="4fe26-143">1.0</span></span>|
|[<span data-ttu-id="4fe26-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4fe26-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4fe26-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4fe26-145">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="4fe26-146">CoercionType: chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-146">CoercionType: String</span></span>

<span data-ttu-id="4fe26-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4fe26-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4fe26-148">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-148">Type</span></span>

*   <span data-ttu-id="4fe26-149">String</span><span class="sxs-lookup"><span data-stu-id="4fe26-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fe26-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4fe26-150">Properties:</span></span>

|<span data-ttu-id="4fe26-151">Nom</span><span class="sxs-lookup"><span data-stu-id="4fe26-151">Name</span></span>| <span data-ttu-id="4fe26-152">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-152">Type</span></span>| <span data-ttu-id="4fe26-153">Description</span><span class="sxs-lookup"><span data-stu-id="4fe26-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4fe26-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-154">String</span></span>|<span data-ttu-id="4fe26-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4fe26-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4fe26-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-156">String</span></span>|<span data-ttu-id="4fe26-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="4fe26-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fe26-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4fe26-158">Requirements</span></span>

|<span data-ttu-id="4fe26-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4fe26-159">Requirement</span></span>| <span data-ttu-id="4fe26-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="4fe26-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fe26-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4fe26-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fe26-162">1.0</span><span class="sxs-lookup"><span data-stu-id="4fe26-162">1.0</span></span>|
|[<span data-ttu-id="4fe26-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4fe26-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4fe26-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4fe26-164">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="4fe26-165">EventType: chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-165">EventType: String</span></span>

<span data-ttu-id="4fe26-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4fe26-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4fe26-167">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-167">Type</span></span>

*   <span data-ttu-id="4fe26-168">String</span><span class="sxs-lookup"><span data-stu-id="4fe26-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fe26-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4fe26-169">Properties:</span></span>

| <span data-ttu-id="4fe26-170">Nom</span><span class="sxs-lookup"><span data-stu-id="4fe26-170">Name</span></span> | <span data-ttu-id="4fe26-171">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-171">Type</span></span> | <span data-ttu-id="4fe26-172">Description</span><span class="sxs-lookup"><span data-stu-id="4fe26-172">Description</span></span> | <span data-ttu-id="4fe26-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="4fe26-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="4fe26-174">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-174">String</span></span> | <span data-ttu-id="4fe26-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="4fe26-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="4fe26-176">1.7</span><span class="sxs-lookup"><span data-stu-id="4fe26-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="4fe26-177">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-177">String</span></span> | <span data-ttu-id="4fe26-178">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="4fe26-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="4fe26-179">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4fe26-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="4fe26-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-180">String</span></span> | <span data-ttu-id="4fe26-181">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="4fe26-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="4fe26-182">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4fe26-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="4fe26-183">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-183">String</span></span> | <span data-ttu-id="4fe26-184">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="4fe26-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="4fe26-185">1,5</span><span class="sxs-lookup"><span data-stu-id="4fe26-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="4fe26-186">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-186">String</span></span> | <span data-ttu-id="4fe26-187">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="4fe26-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="4fe26-188">Aperçu</span><span class="sxs-lookup"><span data-stu-id="4fe26-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="4fe26-189">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-189">String</span></span> | <span data-ttu-id="4fe26-190">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="4fe26-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="4fe26-191">1.7</span><span class="sxs-lookup"><span data-stu-id="4fe26-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="4fe26-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-192">String</span></span> | <span data-ttu-id="4fe26-193">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="4fe26-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="4fe26-194">1.7</span><span class="sxs-lookup"><span data-stu-id="4fe26-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4fe26-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4fe26-195">Requirements</span></span>

|<span data-ttu-id="4fe26-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4fe26-196">Requirement</span></span>| <span data-ttu-id="4fe26-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="4fe26-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fe26-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4fe26-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fe26-199">1,5</span><span class="sxs-lookup"><span data-stu-id="4fe26-199">1.5</span></span> |
|[<span data-ttu-id="4fe26-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4fe26-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4fe26-201">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4fe26-201">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="4fe26-202">SourceProperty: chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-202">SourceProperty: String</span></span>

<span data-ttu-id="4fe26-203">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4fe26-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4fe26-204">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-204">Type</span></span>

*   <span data-ttu-id="4fe26-205">String</span><span class="sxs-lookup"><span data-stu-id="4fe26-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4fe26-206">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4fe26-206">Properties:</span></span>

|<span data-ttu-id="4fe26-207">Nom</span><span class="sxs-lookup"><span data-stu-id="4fe26-207">Name</span></span>| <span data-ttu-id="4fe26-208">Type</span><span class="sxs-lookup"><span data-stu-id="4fe26-208">Type</span></span>| <span data-ttu-id="4fe26-209">Description</span><span class="sxs-lookup"><span data-stu-id="4fe26-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4fe26-210">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-210">String</span></span>|<span data-ttu-id="4fe26-211">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4fe26-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4fe26-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4fe26-212">String</span></span>|<span data-ttu-id="4fe26-213">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4fe26-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4fe26-214">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4fe26-214">Requirements</span></span>

|<span data-ttu-id="4fe26-215">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4fe26-215">Requirement</span></span>| <span data-ttu-id="4fe26-216">Valeur</span><span class="sxs-lookup"><span data-stu-id="4fe26-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="4fe26-217">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4fe26-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4fe26-218">1.0</span><span class="sxs-lookup"><span data-stu-id="4fe26-218">1.0</span></span>|
|[<span data-ttu-id="4fe26-219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4fe26-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4fe26-220">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4fe26-220">Compose or Read</span></span>|
