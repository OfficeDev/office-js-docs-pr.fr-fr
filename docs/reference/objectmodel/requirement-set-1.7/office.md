---
title: Espace de noms Office-ensemble de conditions requises 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 523df189b28fc568ac32e8d17d4a226b52cbd23c
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451618"
---
# <a name="office"></a><span data-ttu-id="e3b6b-102">Office</span><span class="sxs-lookup"><span data-stu-id="e3b6b-102">Office</span></span>

<span data-ttu-id="e3b6b-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e3b6b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e3b6b-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3b6b-105">Requirements</span></span>

|<span data-ttu-id="e3b6b-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3b6b-106">Requirement</span></span>| <span data-ttu-id="e3b6b-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3b6b-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3b6b-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3b6b-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3b6b-109">1.0</span><span class="sxs-lookup"><span data-stu-id="e3b6b-109">1.0</span></span>|
|[<span data-ttu-id="e3b6b-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3b6b-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3b6b-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3b6b-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="e3b6b-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="e3b6b-112">Members and methods</span></span>

| <span data-ttu-id="e3b6b-113">Membre</span><span class="sxs-lookup"><span data-stu-id="e3b6b-113">Member</span></span> | <span data-ttu-id="e3b6b-114">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="e3b6b-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e3b6b-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e3b6b-116">Member</span><span class="sxs-lookup"><span data-stu-id="e3b6b-116">Member</span></span> |
| [<span data-ttu-id="e3b6b-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e3b6b-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e3b6b-118">Member</span><span class="sxs-lookup"><span data-stu-id="e3b6b-118">Member</span></span> |
| [<span data-ttu-id="e3b6b-119">EventType</span><span class="sxs-lookup"><span data-stu-id="e3b6b-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e3b6b-120">Member</span><span class="sxs-lookup"><span data-stu-id="e3b6b-120">Member</span></span> |
| [<span data-ttu-id="e3b6b-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e3b6b-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e3b6b-122">Membre</span><span class="sxs-lookup"><span data-stu-id="e3b6b-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="e3b6b-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="e3b6b-123">Namespaces</span></span>

<span data-ttu-id="e3b6b-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="e3b6b-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-125">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="e3b6b-126">Membres</span><span class="sxs-lookup"><span data-stu-id="e3b6b-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="e3b6b-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="e3b6b-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e3b6b-129">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-129">Type</span></span>

*   <span data-ttu-id="e3b6b-130">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e3b6b-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="e3b6b-131">Properties:</span></span>

|<span data-ttu-id="e3b6b-132">Nom</span><span class="sxs-lookup"><span data-stu-id="e3b6b-132">Name</span></span>| <span data-ttu-id="e3b6b-133">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-133">Type</span></span>| <span data-ttu-id="e3b6b-134">Description</span><span class="sxs-lookup"><span data-stu-id="e3b6b-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e3b6b-135">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-135">String</span></span>|<span data-ttu-id="e3b6b-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e3b6b-137">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-137">String</span></span>|<span data-ttu-id="e3b6b-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3b6b-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3b6b-139">Requirements</span></span>

|<span data-ttu-id="e3b6b-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3b6b-140">Requirement</span></span>| <span data-ttu-id="e3b6b-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3b6b-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3b6b-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3b6b-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3b6b-143">1.0</span><span class="sxs-lookup"><span data-stu-id="e3b6b-143">1.0</span></span>|
|[<span data-ttu-id="e3b6b-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3b6b-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3b6b-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3b6b-145">Compose or Read</span></span>|

---
---

####  <a name="coerciontype-string"></a><span data-ttu-id="e3b6b-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-146">CoercionType :String</span></span>

<span data-ttu-id="e3b6b-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e3b6b-148">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-148">Type</span></span>

*   <span data-ttu-id="e3b6b-149">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e3b6b-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="e3b6b-150">Properties:</span></span>

|<span data-ttu-id="e3b6b-151">Nom</span><span class="sxs-lookup"><span data-stu-id="e3b6b-151">Name</span></span>| <span data-ttu-id="e3b6b-152">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-152">Type</span></span>| <span data-ttu-id="e3b6b-153">Description</span><span class="sxs-lookup"><span data-stu-id="e3b6b-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e3b6b-154">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-154">String</span></span>|<span data-ttu-id="e3b6b-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e3b6b-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3b6b-156">String</span></span>|<span data-ttu-id="e3b6b-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3b6b-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3b6b-158">Requirements</span></span>

|<span data-ttu-id="e3b6b-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3b6b-159">Requirement</span></span>| <span data-ttu-id="e3b6b-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3b6b-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3b6b-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3b6b-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3b6b-162">1.0</span><span class="sxs-lookup"><span data-stu-id="e3b6b-162">1.0</span></span>|
|[<span data-ttu-id="e3b6b-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3b6b-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3b6b-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3b6b-164">Compose or Read</span></span>|

---
---

####  <a name="eventtype-string"></a><span data-ttu-id="e3b6b-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-165">EventType :String</span></span>

<span data-ttu-id="e3b6b-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e3b6b-167">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-167">Type</span></span>

*   <span data-ttu-id="e3b6b-168">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e3b6b-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="e3b6b-169">Properties:</span></span>

| <span data-ttu-id="e3b6b-170">Nom</span><span class="sxs-lookup"><span data-stu-id="e3b6b-170">Name</span></span> | <span data-ttu-id="e3b6b-171">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-171">Type</span></span> | <span data-ttu-id="e3b6b-172">Description</span><span class="sxs-lookup"><span data-stu-id="e3b6b-172">Description</span></span> | <span data-ttu-id="e3b6b-173">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="e3b6b-173">Minimum requirement set</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="e3b6b-174">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-174">String</span></span> | <span data-ttu-id="e3b6b-175">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-175">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e3b6b-176">1.7</span><span class="sxs-lookup"><span data-stu-id="e3b6b-176">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="e3b6b-177">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-177">String</span></span> | <span data-ttu-id="e3b6b-178">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-178">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e3b6b-179">1,5</span><span class="sxs-lookup"><span data-stu-id="e3b6b-179">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e3b6b-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3b6b-180">String</span></span> | <span data-ttu-id="e3b6b-181">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-181">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e3b6b-182">1.7</span><span class="sxs-lookup"><span data-stu-id="e3b6b-182">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e3b6b-183">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e3b6b-183">String</span></span> | <span data-ttu-id="e3b6b-184">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-184">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e3b6b-185">1.7</span><span class="sxs-lookup"><span data-stu-id="e3b6b-185">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e3b6b-186">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3b6b-186">Requirements</span></span>

|<span data-ttu-id="e3b6b-187">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3b6b-187">Requirement</span></span>| <span data-ttu-id="e3b6b-188">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3b6b-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3b6b-189">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3b6b-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3b6b-190">1,5</span><span class="sxs-lookup"><span data-stu-id="e3b6b-190">1.5</span></span> |
|[<span data-ttu-id="e3b6b-191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3b6b-191">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3b6b-192">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3b6b-192">Compose or Read</span></span> |

---
---

####  <a name="sourceproperty-string"></a><span data-ttu-id="e3b6b-193">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-193">SourceProperty :String</span></span>

<span data-ttu-id="e3b6b-194">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-194">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e3b6b-195">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-195">Type</span></span>

*   <span data-ttu-id="e3b6b-196">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-196">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e3b6b-197">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="e3b6b-197">Properties:</span></span>

|<span data-ttu-id="e3b6b-198">Nom</span><span class="sxs-lookup"><span data-stu-id="e3b6b-198">Name</span></span>| <span data-ttu-id="e3b6b-199">Type</span><span class="sxs-lookup"><span data-stu-id="e3b6b-199">Type</span></span>| <span data-ttu-id="e3b6b-200">Description</span><span class="sxs-lookup"><span data-stu-id="e3b6b-200">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e3b6b-201">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-201">String</span></span>|<span data-ttu-id="e3b6b-202">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-202">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e3b6b-203">String</span><span class="sxs-lookup"><span data-stu-id="e3b6b-203">String</span></span>|<span data-ttu-id="e3b6b-204">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="e3b6b-204">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e3b6b-205">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e3b6b-205">Requirements</span></span>

|<span data-ttu-id="e3b6b-206">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e3b6b-206">Requirement</span></span>| <span data-ttu-id="e3b6b-207">Valeur</span><span class="sxs-lookup"><span data-stu-id="e3b6b-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="e3b6b-208">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e3b6b-208">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="e3b6b-209">1.0</span><span class="sxs-lookup"><span data-stu-id="e3b6b-209">1.0</span></span>|
|[<span data-ttu-id="e3b6b-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e3b6b-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="e3b6b-211">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e3b6b-211">Compose or Read</span></span>|
