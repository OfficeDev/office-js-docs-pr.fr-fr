---
title: Espace de noms Office – ensemble de conditions requises
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: df4b47f57d634f6c99ce862ed1c0e96d87be0425
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395665"
---
# <a name="office"></a><span data-ttu-id="22bb0-102">Office</span><span class="sxs-lookup"><span data-stu-id="22bb0-102">Office</span></span>

<span data-ttu-id="22bb0-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="22bb0-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="22bb0-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22bb0-105">Requirements</span></span>

|<span data-ttu-id="22bb0-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22bb0-106">Requirement</span></span>| <span data-ttu-id="22bb0-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="22bb0-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="22bb0-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22bb0-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22bb0-109">1.0</span><span class="sxs-lookup"><span data-stu-id="22bb0-109">1.0</span></span>|
|[<span data-ttu-id="22bb0-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22bb0-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="22bb0-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="22bb0-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="22bb0-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="22bb0-112">Members and methods</span></span>

| <span data-ttu-id="22bb0-113">Membre</span><span class="sxs-lookup"><span data-stu-id="22bb0-113">Member</span></span> | <span data-ttu-id="22bb0-114">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="22bb0-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="22bb0-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="22bb0-116">Member</span><span class="sxs-lookup"><span data-stu-id="22bb0-116">Member</span></span> |
| [<span data-ttu-id="22bb0-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="22bb0-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="22bb0-118">Member</span><span class="sxs-lookup"><span data-stu-id="22bb0-118">Member</span></span> |
| [<span data-ttu-id="22bb0-119">EventType</span><span class="sxs-lookup"><span data-stu-id="22bb0-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="22bb0-120">Member</span><span class="sxs-lookup"><span data-stu-id="22bb0-120">Member</span></span> |
| [<span data-ttu-id="22bb0-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="22bb0-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="22bb0-122">Membre</span><span class="sxs-lookup"><span data-stu-id="22bb0-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="22bb0-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="22bb0-123">Namespaces</span></span>

<span data-ttu-id="22bb0-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="22bb0-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="22bb0-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): inclut un certain nombre d’énumérations, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="22bb0-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="22bb0-126">Members</span><span class="sxs-lookup"><span data-stu-id="22bb0-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="22bb0-127">AsyncResultStatus: chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="22bb0-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="22bb0-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="22bb0-129">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-129">Type</span></span>

*   <span data-ttu-id="22bb0-130">String</span><span class="sxs-lookup"><span data-stu-id="22bb0-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22bb0-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22bb0-131">Properties:</span></span>

|<span data-ttu-id="22bb0-132">Nom</span><span class="sxs-lookup"><span data-stu-id="22bb0-132">Name</span></span>| <span data-ttu-id="22bb0-133">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-133">Type</span></span>| <span data-ttu-id="22bb0-134">Description</span><span class="sxs-lookup"><span data-stu-id="22bb0-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="22bb0-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-135">String</span></span>|<span data-ttu-id="22bb0-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="22bb0-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="22bb0-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-137">String</span></span>|<span data-ttu-id="22bb0-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="22bb0-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22bb0-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22bb0-139">Requirements</span></span>

|<span data-ttu-id="22bb0-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22bb0-140">Requirement</span></span>| <span data-ttu-id="22bb0-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="22bb0-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="22bb0-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22bb0-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22bb0-143">1.0</span><span class="sxs-lookup"><span data-stu-id="22bb0-143">1.0</span></span>|
|[<span data-ttu-id="22bb0-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22bb0-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="22bb0-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="22bb0-145">Compose or Read</span></span>|

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="22bb0-146">CoercionType: chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-146">CoercionType: String</span></span>

<span data-ttu-id="22bb0-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="22bb0-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="22bb0-148">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-148">Type</span></span>

*   <span data-ttu-id="22bb0-149">String</span><span class="sxs-lookup"><span data-stu-id="22bb0-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22bb0-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22bb0-150">Properties:</span></span>

|<span data-ttu-id="22bb0-151">Nom</span><span class="sxs-lookup"><span data-stu-id="22bb0-151">Name</span></span>| <span data-ttu-id="22bb0-152">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-152">Type</span></span>| <span data-ttu-id="22bb0-153">Description</span><span class="sxs-lookup"><span data-stu-id="22bb0-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="22bb0-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-154">String</span></span>|<span data-ttu-id="22bb0-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="22bb0-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="22bb0-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-156">String</span></span>|<span data-ttu-id="22bb0-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="22bb0-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22bb0-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22bb0-158">Requirements</span></span>

|<span data-ttu-id="22bb0-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22bb0-159">Requirement</span></span>| <span data-ttu-id="22bb0-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="22bb0-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="22bb0-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22bb0-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22bb0-162">1.0</span><span class="sxs-lookup"><span data-stu-id="22bb0-162">1.0</span></span>|
|[<span data-ttu-id="22bb0-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22bb0-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="22bb0-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="22bb0-164">Compose or Read</span></span>|

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="22bb0-165">EventType: chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-165">EventType: String</span></span>

<span data-ttu-id="22bb0-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="22bb0-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="22bb0-167">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-167">Type</span></span>

*   <span data-ttu-id="22bb0-168">String</span><span class="sxs-lookup"><span data-stu-id="22bb0-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22bb0-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22bb0-169">Properties:</span></span>

| <span data-ttu-id="22bb0-170">Nom</span><span class="sxs-lookup"><span data-stu-id="22bb0-170">Name</span></span> | <span data-ttu-id="22bb0-171">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-171">Type</span></span> | <span data-ttu-id="22bb0-172">Description</span><span class="sxs-lookup"><span data-stu-id="22bb0-172">Description</span></span> | <span data-ttu-id="22bb0-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="22bb0-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="22bb0-174">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-174">String</span></span> | <span data-ttu-id="22bb0-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="22bb0-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="22bb0-176">1.7</span><span class="sxs-lookup"><span data-stu-id="22bb0-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="22bb0-177">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-177">String</span></span> | <span data-ttu-id="22bb0-178">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="22bb0-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="22bb0-179">Aperçu</span><span class="sxs-lookup"><span data-stu-id="22bb0-179">Preview</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="22bb0-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-180">String</span></span> | <span data-ttu-id="22bb0-181">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="22bb0-181">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="22bb0-182">Aperçu</span><span class="sxs-lookup"><span data-stu-id="22bb0-182">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="22bb0-183">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-183">String</span></span> | <span data-ttu-id="22bb0-184">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="22bb0-184">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="22bb0-185">1,5</span><span class="sxs-lookup"><span data-stu-id="22bb0-185">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="22bb0-186">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-186">String</span></span> | <span data-ttu-id="22bb0-187">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="22bb0-187">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="22bb0-188">Aperçu</span><span class="sxs-lookup"><span data-stu-id="22bb0-188">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="22bb0-189">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-189">String</span></span> | <span data-ttu-id="22bb0-190">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="22bb0-190">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="22bb0-191">1.7</span><span class="sxs-lookup"><span data-stu-id="22bb0-191">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="22bb0-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-192">String</span></span> | <span data-ttu-id="22bb0-193">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="22bb0-193">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="22bb0-194">1.7</span><span class="sxs-lookup"><span data-stu-id="22bb0-194">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="22bb0-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22bb0-195">Requirements</span></span>

|<span data-ttu-id="22bb0-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22bb0-196">Requirement</span></span>| <span data-ttu-id="22bb0-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="22bb0-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="22bb0-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22bb0-198">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22bb0-199">1,5</span><span class="sxs-lookup"><span data-stu-id="22bb0-199">1.5</span></span> |
|[<span data-ttu-id="22bb0-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22bb0-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="22bb0-201">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="22bb0-201">Compose or Read</span></span> |

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="22bb0-202">SourceProperty: chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-202">SourceProperty: String</span></span>

<span data-ttu-id="22bb0-203">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="22bb0-203">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="22bb0-204">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-204">Type</span></span>

*   <span data-ttu-id="22bb0-205">String</span><span class="sxs-lookup"><span data-stu-id="22bb0-205">String</span></span>

##### <a name="properties"></a><span data-ttu-id="22bb0-206">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="22bb0-206">Properties:</span></span>

|<span data-ttu-id="22bb0-207">Nom</span><span class="sxs-lookup"><span data-stu-id="22bb0-207">Name</span></span>| <span data-ttu-id="22bb0-208">Type</span><span class="sxs-lookup"><span data-stu-id="22bb0-208">Type</span></span>| <span data-ttu-id="22bb0-209">Description</span><span class="sxs-lookup"><span data-stu-id="22bb0-209">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="22bb0-210">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-210">String</span></span>|<span data-ttu-id="22bb0-211">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="22bb0-211">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="22bb0-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22bb0-212">String</span></span>|<span data-ttu-id="22bb0-213">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="22bb0-213">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="22bb0-214">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22bb0-214">Requirements</span></span>

|<span data-ttu-id="22bb0-215">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22bb0-215">Requirement</span></span>| <span data-ttu-id="22bb0-216">Valeur</span><span class="sxs-lookup"><span data-stu-id="22bb0-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="22bb0-217">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22bb0-217">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22bb0-218">1.0</span><span class="sxs-lookup"><span data-stu-id="22bb0-218">1.0</span></span>|
|[<span data-ttu-id="22bb0-219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22bb0-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="22bb0-220">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="22bb0-220">Compose or Read</span></span>|
