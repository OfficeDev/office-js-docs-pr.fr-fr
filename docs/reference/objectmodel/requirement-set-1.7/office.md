---
title: Espace de noms Office – ensemble de conditions requises 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: d6422e470864d5a02db37e1fef295e8cbb82a213
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067894"
---
# <a name="office"></a><span data-ttu-id="0a19d-102">Office</span><span class="sxs-lookup"><span data-stu-id="0a19d-102">Office</span></span>

<span data-ttu-id="0a19d-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0a19d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0a19d-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0a19d-105">Requirements</span></span>

|<span data-ttu-id="0a19d-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0a19d-106">Requirement</span></span>| <span data-ttu-id="0a19d-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="0a19d-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a19d-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0a19d-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a19d-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0a19d-109">1.0</span></span>|
|[<span data-ttu-id="0a19d-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0a19d-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a19d-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0a19d-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0a19d-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="0a19d-112">Members and methods</span></span>

| <span data-ttu-id="0a19d-113">Membre</span><span class="sxs-lookup"><span data-stu-id="0a19d-113">Member</span></span> | <span data-ttu-id="0a19d-114">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0a19d-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0a19d-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0a19d-116">Membre</span><span class="sxs-lookup"><span data-stu-id="0a19d-116">Member</span></span> |
| [<span data-ttu-id="0a19d-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0a19d-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0a19d-118">Membre</span><span class="sxs-lookup"><span data-stu-id="0a19d-118">Member</span></span> |
| [<span data-ttu-id="0a19d-119">EventType</span><span class="sxs-lookup"><span data-stu-id="0a19d-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0a19d-120">Membre</span><span class="sxs-lookup"><span data-stu-id="0a19d-120">Member</span></span> |
| [<span data-ttu-id="0a19d-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0a19d-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0a19d-122">Membre</span><span class="sxs-lookup"><span data-stu-id="0a19d-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0a19d-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="0a19d-123">Namespaces</span></span>

<span data-ttu-id="0a19d-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="0a19d-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0a19d-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="0a19d-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="0a19d-126">Membres</span><span class="sxs-lookup"><span data-stu-id="0a19d-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="0a19d-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="0a19d-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="0a19d-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0a19d-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0a19d-129">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-129">Type</span></span>

*   <span data-ttu-id="0a19d-130">String</span><span class="sxs-lookup"><span data-stu-id="0a19d-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0a19d-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0a19d-131">Properties:</span></span>

|<span data-ttu-id="0a19d-132">Nom</span><span class="sxs-lookup"><span data-stu-id="0a19d-132">Name</span></span>| <span data-ttu-id="0a19d-133">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-133">Type</span></span>| <span data-ttu-id="0a19d-134">Description</span><span class="sxs-lookup"><span data-stu-id="0a19d-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0a19d-135">String</span><span class="sxs-lookup"><span data-stu-id="0a19d-135">String</span></span>|<span data-ttu-id="0a19d-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="0a19d-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0a19d-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-137">String</span></span>|<span data-ttu-id="0a19d-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="0a19d-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0a19d-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0a19d-139">Requirements</span></span>

|<span data-ttu-id="0a19d-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0a19d-140">Requirement</span></span>| <span data-ttu-id="0a19d-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="0a19d-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a19d-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0a19d-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a19d-143">1.0</span><span class="sxs-lookup"><span data-stu-id="0a19d-143">1.0</span></span>|
|[<span data-ttu-id="0a19d-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0a19d-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a19d-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0a19d-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="0a19d-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="0a19d-146">CoercionType :String</span></span>

<span data-ttu-id="0a19d-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="0a19d-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0a19d-148">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-148">Type</span></span>

*   <span data-ttu-id="0a19d-149">String</span><span class="sxs-lookup"><span data-stu-id="0a19d-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0a19d-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0a19d-150">Properties:</span></span>

|<span data-ttu-id="0a19d-151">Nom</span><span class="sxs-lookup"><span data-stu-id="0a19d-151">Name</span></span>| <span data-ttu-id="0a19d-152">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-152">Type</span></span>| <span data-ttu-id="0a19d-153">Description</span><span class="sxs-lookup"><span data-stu-id="0a19d-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0a19d-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-154">String</span></span>|<span data-ttu-id="0a19d-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="0a19d-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0a19d-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-156">String</span></span>|<span data-ttu-id="0a19d-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="0a19d-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0a19d-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0a19d-158">Requirements</span></span>

|<span data-ttu-id="0a19d-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0a19d-159">Requirement</span></span>| <span data-ttu-id="0a19d-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="0a19d-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a19d-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0a19d-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a19d-162">1.0</span><span class="sxs-lookup"><span data-stu-id="0a19d-162">1.0</span></span>|
|[<span data-ttu-id="0a19d-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0a19d-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a19d-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0a19d-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="0a19d-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="0a19d-165">EventType :String</span></span>

<span data-ttu-id="0a19d-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="0a19d-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0a19d-167">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-167">Type</span></span>

*   <span data-ttu-id="0a19d-168">String</span><span class="sxs-lookup"><span data-stu-id="0a19d-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0a19d-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0a19d-169">Properties:</span></span>

| <span data-ttu-id="0a19d-170">Nom</span><span class="sxs-lookup"><span data-stu-id="0a19d-170">Name</span></span> | <span data-ttu-id="0a19d-171">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-171">Type</span></span> | <span data-ttu-id="0a19d-172">Description</span><span class="sxs-lookup"><span data-stu-id="0a19d-172">Description</span></span> | <span data-ttu-id="0a19d-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="0a19d-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="0a19d-174">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-174">String</span></span> | <span data-ttu-id="0a19d-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="0a19d-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="0a19d-176">1.7</span><span class="sxs-lookup"><span data-stu-id="0a19d-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="0a19d-177">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-177">String</span></span> | <span data-ttu-id="0a19d-178">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="0a19d-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="0a19d-179">1,5</span><span class="sxs-lookup"><span data-stu-id="0a19d-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="0a19d-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-180">String</span></span> | <span data-ttu-id="0a19d-181">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="0a19d-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="0a19d-182">1.7</span><span class="sxs-lookup"><span data-stu-id="0a19d-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="0a19d-183">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-183">String</span></span> | <span data-ttu-id="0a19d-184">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="0a19d-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="0a19d-185">1.7</span><span class="sxs-lookup"><span data-stu-id="0a19d-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0a19d-186">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0a19d-186">Requirements</span></span>

|<span data-ttu-id="0a19d-187">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0a19d-187">Requirement</span></span>| <span data-ttu-id="0a19d-188">Valeur</span><span class="sxs-lookup"><span data-stu-id="0a19d-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a19d-189">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0a19d-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a19d-190">1,5</span><span class="sxs-lookup"><span data-stu-id="0a19d-190">1.5</span></span> |
|[<span data-ttu-id="0a19d-191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0a19d-191">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a19d-192">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0a19d-192">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="0a19d-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="0a19d-193">SourceProperty :String</span></span>

<span data-ttu-id="0a19d-194">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="0a19d-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0a19d-195">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-195">Type</span></span>

*   <span data-ttu-id="0a19d-196">String</span><span class="sxs-lookup"><span data-stu-id="0a19d-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0a19d-197">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0a19d-197">Properties:</span></span>

|<span data-ttu-id="0a19d-198">Nom</span><span class="sxs-lookup"><span data-stu-id="0a19d-198">Name</span></span>| <span data-ttu-id="0a19d-199">Type</span><span class="sxs-lookup"><span data-stu-id="0a19d-199">Type</span></span>| <span data-ttu-id="0a19d-200">Description</span><span class="sxs-lookup"><span data-stu-id="0a19d-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0a19d-201">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0a19d-201">String</span></span>|<span data-ttu-id="0a19d-202">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="0a19d-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0a19d-203">String</span><span class="sxs-lookup"><span data-stu-id="0a19d-203">String</span></span>|<span data-ttu-id="0a19d-204">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="0a19d-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0a19d-205">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0a19d-205">Requirements</span></span>

|<span data-ttu-id="0a19d-206">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0a19d-206">Requirement</span></span>| <span data-ttu-id="0a19d-207">Valeur</span><span class="sxs-lookup"><span data-stu-id="0a19d-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a19d-208">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0a19d-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a19d-209">1.0</span><span class="sxs-lookup"><span data-stu-id="0a19d-209">1.0</span></span>|
|[<span data-ttu-id="0a19d-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0a19d-210">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a19d-211">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="0a19d-211">Compose or Read</span></span>|
