---
title: Espace de noms Office – ensemble de conditions requises
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: f4a4f0d7a4ce0de433d4e70b6a4675b5f63f26f0
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457928"
---
# <a name="office"></a><span data-ttu-id="2cdd7-102">Office</span><span class="sxs-lookup"><span data-stu-id="2cdd7-102">Office</span></span>

<span data-ttu-id="2cdd7-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="2cdd7-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2cdd7-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2cdd7-105">Requirements</span></span>

|<span data-ttu-id="2cdd7-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2cdd7-106">Requirement</span></span>| <span data-ttu-id="2cdd7-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="2cdd7-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2cdd7-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2cdd7-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2cdd7-109">1.0</span><span class="sxs-lookup"><span data-stu-id="2cdd7-109">1.0</span></span>|
|[<span data-ttu-id="2cdd7-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2cdd7-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2cdd7-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2cdd7-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2cdd7-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="2cdd7-112">Members and methods</span></span>

| <span data-ttu-id="2cdd7-113">Membre</span><span class="sxs-lookup"><span data-stu-id="2cdd7-113">Member</span></span> | <span data-ttu-id="2cdd7-114">Type</span><span class="sxs-lookup"><span data-stu-id="2cdd7-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2cdd7-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2cdd7-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2cdd7-116">Membre</span><span class="sxs-lookup"><span data-stu-id="2cdd7-116">Member</span></span> |
| [<span data-ttu-id="2cdd7-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2cdd7-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2cdd7-118">Membre</span><span class="sxs-lookup"><span data-stu-id="2cdd7-118">Member</span></span> |
| [<span data-ttu-id="2cdd7-119">EventType</span><span class="sxs-lookup"><span data-stu-id="2cdd7-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2cdd7-120">Membre</span><span class="sxs-lookup"><span data-stu-id="2cdd7-120">Member</span></span> |
| [<span data-ttu-id="2cdd7-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2cdd7-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2cdd7-122">Membre</span><span class="sxs-lookup"><span data-stu-id="2cdd7-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="2cdd7-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="2cdd7-123">Namespaces</span></span>

<span data-ttu-id="2cdd7-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="2cdd7-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="2cdd7-126">Membres</span><span class="sxs-lookup"><span data-stu-id="2cdd7-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="2cdd7-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="2cdd7-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2cdd7-129">Type :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-129">Type:</span></span>

*   <span data-ttu-id="2cdd7-130">String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2cdd7-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-131">Properties:</span></span>

|<span data-ttu-id="2cdd7-132">Nom</span><span class="sxs-lookup"><span data-stu-id="2cdd7-132">Name</span></span>| <span data-ttu-id="2cdd7-133">Type</span><span class="sxs-lookup"><span data-stu-id="2cdd7-133">Type</span></span>| <span data-ttu-id="2cdd7-134">Description</span><span class="sxs-lookup"><span data-stu-id="2cdd7-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2cdd7-135">String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-135">String</span></span>|<span data-ttu-id="2cdd7-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2cdd7-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-137">String</span></span>|<span data-ttu-id="2cdd7-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2cdd7-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2cdd7-139">Requirements</span></span>

|<span data-ttu-id="2cdd7-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2cdd7-140">Requirement</span></span>| <span data-ttu-id="2cdd7-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="2cdd7-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="2cdd7-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2cdd7-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2cdd7-143">1.0</span><span class="sxs-lookup"><span data-stu-id="2cdd7-143">1.0</span></span>|
|[<span data-ttu-id="2cdd7-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2cdd7-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2cdd7-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2cdd7-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="2cdd7-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-146">CoercionType :String</span></span>

<span data-ttu-id="2cdd7-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2cdd7-148">Type :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-148">Type:</span></span>

*   <span data-ttu-id="2cdd7-149">String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2cdd7-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-150">Properties:</span></span>

|<span data-ttu-id="2cdd7-151">Nom</span><span class="sxs-lookup"><span data-stu-id="2cdd7-151">Name</span></span>| <span data-ttu-id="2cdd7-152">Type</span><span class="sxs-lookup"><span data-stu-id="2cdd7-152">Type</span></span>| <span data-ttu-id="2cdd7-153">Description</span><span class="sxs-lookup"><span data-stu-id="2cdd7-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2cdd7-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-154">String</span></span>|<span data-ttu-id="2cdd7-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2cdd7-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-156">String</span></span>|<span data-ttu-id="2cdd7-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2cdd7-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2cdd7-158">Requirements</span></span>

|<span data-ttu-id="2cdd7-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2cdd7-159">Requirement</span></span>| <span data-ttu-id="2cdd7-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="2cdd7-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="2cdd7-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2cdd7-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2cdd7-162">1.0</span><span class="sxs-lookup"><span data-stu-id="2cdd7-162">1.0</span></span>|
|[<span data-ttu-id="2cdd7-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2cdd7-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2cdd7-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2cdd7-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="2cdd7-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-165">EventType :String</span></span>

<span data-ttu-id="2cdd7-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2cdd7-167">Type :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-167">Type:</span></span>

*   <span data-ttu-id="2cdd7-168">String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2cdd7-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-169">Properties:</span></span>

| <span data-ttu-id="2cdd7-170">Nom</span><span class="sxs-lookup"><span data-stu-id="2cdd7-170">Name</span></span> | <span data-ttu-id="2cdd7-171">Type</span><span class="sxs-lookup"><span data-stu-id="2cdd7-171">Type</span></span> | <span data-ttu-id="2cdd7-172">Description</span><span class="sxs-lookup"><span data-stu-id="2cdd7-172">Description</span></span> | <span data-ttu-id="2cdd7-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="2cdd7-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="2cdd7-174">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-174">String</span></span> | <span data-ttu-id="2cdd7-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="2cdd7-176">1.7</span><span class="sxs-lookup"><span data-stu-id="2cdd7-176">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="2cdd7-177">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-177">String</span></span> | <span data-ttu-id="2cdd7-178">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-178">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="2cdd7-179">Aperçu</span><span class="sxs-lookup"><span data-stu-id="2cdd7-179">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="2cdd7-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-180">String</span></span> | <span data-ttu-id="2cdd7-181">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-181">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2cdd7-182">1,5</span><span class="sxs-lookup"><span data-stu-id="2cdd7-182">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="2cdd7-183">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-183">String</span></span> | <span data-ttu-id="2cdd7-184">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-184">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="2cdd7-185">Aperçu</span><span class="sxs-lookup"><span data-stu-id="2cdd7-185">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="2cdd7-186">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-186">String</span></span> | <span data-ttu-id="2cdd7-187">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-187">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="2cdd7-188">1.7</span><span class="sxs-lookup"><span data-stu-id="2cdd7-188">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="2cdd7-189">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-189">String</span></span> | <span data-ttu-id="2cdd7-190">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-190">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="2cdd7-191">1.7</span><span class="sxs-lookup"><span data-stu-id="2cdd7-191">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2cdd7-192">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2cdd7-192">Requirements</span></span>

|<span data-ttu-id="2cdd7-193">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2cdd7-193">Requirement</span></span>| <span data-ttu-id="2cdd7-194">Valeur</span><span class="sxs-lookup"><span data-stu-id="2cdd7-194">Value</span></span>|
|---|---|
|[<span data-ttu-id="2cdd7-195">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2cdd7-195">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2cdd7-196">1,5</span><span class="sxs-lookup"><span data-stu-id="2cdd7-196">1.5</span></span> |
|[<span data-ttu-id="2cdd7-197">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2cdd7-197">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2cdd7-198">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2cdd7-198">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="2cdd7-199">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-199">SourceProperty :String</span></span>

<span data-ttu-id="2cdd7-200">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-200">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2cdd7-201">Type :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-201">Type:</span></span>

*   <span data-ttu-id="2cdd7-202">String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-202">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2cdd7-203">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2cdd7-203">Properties:</span></span>

|<span data-ttu-id="2cdd7-204">Nom</span><span class="sxs-lookup"><span data-stu-id="2cdd7-204">Name</span></span>| <span data-ttu-id="2cdd7-205">Type</span><span class="sxs-lookup"><span data-stu-id="2cdd7-205">Type</span></span>| <span data-ttu-id="2cdd7-206">Description</span><span class="sxs-lookup"><span data-stu-id="2cdd7-206">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2cdd7-207">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2cdd7-207">String</span></span>|<span data-ttu-id="2cdd7-208">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-208">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2cdd7-209">String</span><span class="sxs-lookup"><span data-stu-id="2cdd7-209">String</span></span>|<span data-ttu-id="2cdd7-210">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="2cdd7-210">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2cdd7-211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2cdd7-211">Requirements</span></span>

|<span data-ttu-id="2cdd7-212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2cdd7-212">Requirement</span></span>| <span data-ttu-id="2cdd7-213">Valeur</span><span class="sxs-lookup"><span data-stu-id="2cdd7-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="2cdd7-214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2cdd7-214">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2cdd7-215">1.0</span><span class="sxs-lookup"><span data-stu-id="2cdd7-215">1.0</span></span>|
|[<span data-ttu-id="2cdd7-216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2cdd7-216">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2cdd7-217">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="2cdd7-217">Compose or read</span></span>|