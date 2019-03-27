---
title: Espace de noms Office-ensemble de conditions requises 1,7
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: 533e997fc7f8be6eb6d3aefefaf023e8c7666af2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870526"
---
# <a name="office"></a><span data-ttu-id="b78b1-102">Office</span><span class="sxs-lookup"><span data-stu-id="b78b1-102">Office</span></span>

<span data-ttu-id="b78b1-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b78b1-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b78b1-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b78b1-105">Requirements</span></span>

|<span data-ttu-id="b78b1-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b78b1-106">Requirement</span></span>| <span data-ttu-id="b78b1-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="b78b1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b78b1-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b78b1-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b78b1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b78b1-109">1.0</span></span>|
|[<span data-ttu-id="b78b1-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b78b1-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b78b1-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b78b1-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b78b1-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="b78b1-112">Members and methods</span></span>

| <span data-ttu-id="b78b1-113">Membre</span><span class="sxs-lookup"><span data-stu-id="b78b1-113">Member</span></span> | <span data-ttu-id="b78b1-114">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b78b1-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b78b1-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b78b1-116">Member</span><span class="sxs-lookup"><span data-stu-id="b78b1-116">Member</span></span> |
| [<span data-ttu-id="b78b1-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b78b1-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b78b1-118">Member</span><span class="sxs-lookup"><span data-stu-id="b78b1-118">Member</span></span> |
| [<span data-ttu-id="b78b1-119">EventType</span><span class="sxs-lookup"><span data-stu-id="b78b1-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="b78b1-120">Member</span><span class="sxs-lookup"><span data-stu-id="b78b1-120">Member</span></span> |
| [<span data-ttu-id="b78b1-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b78b1-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b78b1-122">Membre</span><span class="sxs-lookup"><span data-stu-id="b78b1-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b78b1-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="b78b1-123">Namespaces</span></span>

<span data-ttu-id="b78b1-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="b78b1-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="b78b1-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="b78b1-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="b78b1-126">Membres</span><span class="sxs-lookup"><span data-stu-id="b78b1-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="b78b1-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="b78b1-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="b78b1-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b78b1-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b78b1-129">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-129">Type</span></span>

*   <span data-ttu-id="b78b1-130">String</span><span class="sxs-lookup"><span data-stu-id="b78b1-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b78b1-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b78b1-131">Properties:</span></span>

|<span data-ttu-id="b78b1-132">Nom</span><span class="sxs-lookup"><span data-stu-id="b78b1-132">Name</span></span>| <span data-ttu-id="b78b1-133">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-133">Type</span></span>| <span data-ttu-id="b78b1-134">Description</span><span class="sxs-lookup"><span data-stu-id="b78b1-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b78b1-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-135">String</span></span>|<span data-ttu-id="b78b1-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="b78b1-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b78b1-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-137">String</span></span>|<span data-ttu-id="b78b1-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="b78b1-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b78b1-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b78b1-139">Requirements</span></span>

|<span data-ttu-id="b78b1-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b78b1-140">Requirement</span></span>| <span data-ttu-id="b78b1-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="b78b1-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="b78b1-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b78b1-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b78b1-143">1.0</span><span class="sxs-lookup"><span data-stu-id="b78b1-143">1.0</span></span>|
|[<span data-ttu-id="b78b1-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b78b1-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b78b1-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b78b1-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="b78b1-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="b78b1-146">CoercionType :String</span></span>

<span data-ttu-id="b78b1-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="b78b1-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b78b1-148">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-148">Type</span></span>

*   <span data-ttu-id="b78b1-149">String</span><span class="sxs-lookup"><span data-stu-id="b78b1-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b78b1-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b78b1-150">Properties:</span></span>

|<span data-ttu-id="b78b1-151">Nom</span><span class="sxs-lookup"><span data-stu-id="b78b1-151">Name</span></span>| <span data-ttu-id="b78b1-152">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-152">Type</span></span>| <span data-ttu-id="b78b1-153">Description</span><span class="sxs-lookup"><span data-stu-id="b78b1-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b78b1-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-154">String</span></span>|<span data-ttu-id="b78b1-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="b78b1-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b78b1-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-156">String</span></span>|<span data-ttu-id="b78b1-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="b78b1-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b78b1-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b78b1-158">Requirements</span></span>

|<span data-ttu-id="b78b1-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b78b1-159">Requirement</span></span>| <span data-ttu-id="b78b1-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="b78b1-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="b78b1-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b78b1-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b78b1-162">1.0</span><span class="sxs-lookup"><span data-stu-id="b78b1-162">1.0</span></span>|
|[<span data-ttu-id="b78b1-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b78b1-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b78b1-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b78b1-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="b78b1-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="b78b1-165">EventType :String</span></span>

<span data-ttu-id="b78b1-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="b78b1-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="b78b1-167">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-167">Type</span></span>

*   <span data-ttu-id="b78b1-168">String</span><span class="sxs-lookup"><span data-stu-id="b78b1-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b78b1-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b78b1-169">Properties:</span></span>

| <span data-ttu-id="b78b1-170">Nom</span><span class="sxs-lookup"><span data-stu-id="b78b1-170">Name</span></span> | <span data-ttu-id="b78b1-171">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-171">Type</span></span> | <span data-ttu-id="b78b1-172">Description</span><span class="sxs-lookup"><span data-stu-id="b78b1-172">Description</span></span> | <span data-ttu-id="b78b1-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="b78b1-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="b78b1-174">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-174">String</span></span> | <span data-ttu-id="b78b1-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="b78b1-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="b78b1-176">1.7</span><span class="sxs-lookup"><span data-stu-id="b78b1-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="b78b1-177">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-177">String</span></span> | <span data-ttu-id="b78b1-178">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="b78b1-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="b78b1-179">1,5</span><span class="sxs-lookup"><span data-stu-id="b78b1-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="b78b1-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-180">String</span></span> | <span data-ttu-id="b78b1-181">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="b78b1-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="b78b1-182">1.7</span><span class="sxs-lookup"><span data-stu-id="b78b1-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="b78b1-183">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-183">String</span></span> | <span data-ttu-id="b78b1-184">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="b78b1-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="b78b1-185">1.7</span><span class="sxs-lookup"><span data-stu-id="b78b1-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b78b1-186">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b78b1-186">Requirements</span></span>

|<span data-ttu-id="b78b1-187">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b78b1-187">Requirement</span></span>| <span data-ttu-id="b78b1-188">Valeur</span><span class="sxs-lookup"><span data-stu-id="b78b1-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="b78b1-189">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b78b1-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b78b1-190">1,5</span><span class="sxs-lookup"><span data-stu-id="b78b1-190">1.5</span></span> |
|[<span data-ttu-id="b78b1-191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b78b1-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b78b1-192">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b78b1-192">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="b78b1-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="b78b1-193">SourceProperty :String</span></span>

<span data-ttu-id="b78b1-194">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="b78b1-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b78b1-195">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-195">Type</span></span>

*   <span data-ttu-id="b78b1-196">String</span><span class="sxs-lookup"><span data-stu-id="b78b1-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b78b1-197">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b78b1-197">Properties:</span></span>

|<span data-ttu-id="b78b1-198">Nom</span><span class="sxs-lookup"><span data-stu-id="b78b1-198">Name</span></span>| <span data-ttu-id="b78b1-199">Type</span><span class="sxs-lookup"><span data-stu-id="b78b1-199">Type</span></span>| <span data-ttu-id="b78b1-200">Description</span><span class="sxs-lookup"><span data-stu-id="b78b1-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b78b1-201">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-201">String</span></span>|<span data-ttu-id="b78b1-202">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="b78b1-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b78b1-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b78b1-203">String</span></span>|<span data-ttu-id="b78b1-204">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="b78b1-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b78b1-205">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b78b1-205">Requirements</span></span>

|<span data-ttu-id="b78b1-206">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b78b1-206">Requirement</span></span>| <span data-ttu-id="b78b1-207">Valeur</span><span class="sxs-lookup"><span data-stu-id="b78b1-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="b78b1-208">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b78b1-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b78b1-209">1.0</span><span class="sxs-lookup"><span data-stu-id="b78b1-209">1.0</span></span>|
|[<span data-ttu-id="b78b1-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b78b1-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b78b1-211">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b78b1-211">Compose or Read</span></span>|
