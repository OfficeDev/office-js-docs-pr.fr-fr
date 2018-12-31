---
title: Espace de noms Office-ensemble de conditions requises 1.5
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 11b9ea439e659f0aefdcd15ae9a73ac128aee98b
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27458005"
---
# <a name="office"></a><span data-ttu-id="4ffd2-102">Office</span><span class="sxs-lookup"><span data-stu-id="4ffd2-102">Office</span></span>

<span data-ttu-id="4ffd2-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="4ffd2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ffd2-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ffd2-105">Requirements</span></span>

|<span data-ttu-id="4ffd2-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ffd2-106">Requirement</span></span>| <span data-ttu-id="4ffd2-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ffd2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ffd2-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ffd2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ffd2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4ffd2-109">1.0</span></span>|
|[<span data-ttu-id="4ffd2-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ffd2-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ffd2-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ffd2-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4ffd2-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4ffd2-112">Members and methods</span></span>

| <span data-ttu-id="4ffd2-113">Membre</span><span class="sxs-lookup"><span data-stu-id="4ffd2-113">Member</span></span> | <span data-ttu-id="4ffd2-114">Type</span><span class="sxs-lookup"><span data-stu-id="4ffd2-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4ffd2-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="4ffd2-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="4ffd2-116">Membre</span><span class="sxs-lookup"><span data-stu-id="4ffd2-116">Member</span></span> |
| [<span data-ttu-id="4ffd2-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="4ffd2-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="4ffd2-118">Membre</span><span class="sxs-lookup"><span data-stu-id="4ffd2-118">Member</span></span> |
| [<span data-ttu-id="4ffd2-119">EventType</span><span class="sxs-lookup"><span data-stu-id="4ffd2-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="4ffd2-120">Membre</span><span class="sxs-lookup"><span data-stu-id="4ffd2-120">Member</span></span> |
| [<span data-ttu-id="4ffd2-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="4ffd2-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="4ffd2-122">Membre</span><span class="sxs-lookup"><span data-stu-id="4ffd2-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4ffd2-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4ffd2-123">Namespaces</span></span>

<span data-ttu-id="4ffd2-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="4ffd2-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="4ffd2-126">Membres</span><span class="sxs-lookup"><span data-stu-id="4ffd2-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="4ffd2-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="4ffd2-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="4ffd2-129">Type :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-129">Type:</span></span>

*   <span data-ttu-id="4ffd2-130">String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ffd2-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-131">Properties:</span></span>

|<span data-ttu-id="4ffd2-132">Nom</span><span class="sxs-lookup"><span data-stu-id="4ffd2-132">Name</span></span>| <span data-ttu-id="4ffd2-133">Type</span><span class="sxs-lookup"><span data-stu-id="4ffd2-133">Type</span></span>| <span data-ttu-id="4ffd2-134">Description</span><span class="sxs-lookup"><span data-stu-id="4ffd2-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="4ffd2-135">String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-135">String</span></span>|<span data-ttu-id="4ffd2-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="4ffd2-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4ffd2-137">String</span></span>|<span data-ttu-id="4ffd2-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ffd2-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ffd2-139">Requirements</span></span>

|<span data-ttu-id="4ffd2-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ffd2-140">Requirement</span></span>| <span data-ttu-id="4ffd2-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ffd2-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ffd2-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ffd2-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ffd2-143">1.0</span><span class="sxs-lookup"><span data-stu-id="4ffd2-143">1.0</span></span>|
|[<span data-ttu-id="4ffd2-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ffd2-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ffd2-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ffd2-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="4ffd2-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-146">CoercionType :String</span></span>

<span data-ttu-id="4ffd2-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4ffd2-148">Type :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-148">Type:</span></span>

*   <span data-ttu-id="4ffd2-149">String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ffd2-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-150">Properties:</span></span>

|<span data-ttu-id="4ffd2-151">Nom</span><span class="sxs-lookup"><span data-stu-id="4ffd2-151">Name</span></span>| <span data-ttu-id="4ffd2-152">Type</span><span class="sxs-lookup"><span data-stu-id="4ffd2-152">Type</span></span>| <span data-ttu-id="4ffd2-153">Description</span><span class="sxs-lookup"><span data-stu-id="4ffd2-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="4ffd2-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4ffd2-154">String</span></span>|<span data-ttu-id="4ffd2-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="4ffd2-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4ffd2-156">String</span></span>|<span data-ttu-id="4ffd2-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ffd2-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ffd2-158">Requirements</span></span>

|<span data-ttu-id="4ffd2-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ffd2-159">Requirement</span></span>| <span data-ttu-id="4ffd2-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ffd2-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ffd2-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ffd2-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ffd2-162">1.0</span><span class="sxs-lookup"><span data-stu-id="4ffd2-162">1.0</span></span>|
|[<span data-ttu-id="4ffd2-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ffd2-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ffd2-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ffd2-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="4ffd2-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-165">EventType :String</span></span>

<span data-ttu-id="4ffd2-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="4ffd2-167">Type :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-167">Type:</span></span>

*   <span data-ttu-id="4ffd2-168">String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ffd2-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-169">Properties:</span></span>

| <span data-ttu-id="4ffd2-170">Nom</span><span class="sxs-lookup"><span data-stu-id="4ffd2-170">Name</span></span> | <span data-ttu-id="4ffd2-171">Type</span><span class="sxs-lookup"><span data-stu-id="4ffd2-171">Type</span></span> | <span data-ttu-id="4ffd2-172">Description</span><span class="sxs-lookup"><span data-stu-id="4ffd2-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="4ffd2-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4ffd2-173">String</span></span> | <span data-ttu-id="4ffd2-174">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4ffd2-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ffd2-175">Requirements</span></span>

|<span data-ttu-id="4ffd2-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ffd2-176">Requirement</span></span>| <span data-ttu-id="4ffd2-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ffd2-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ffd2-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ffd2-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ffd2-179">1,5</span><span class="sxs-lookup"><span data-stu-id="4ffd2-179">1.5</span></span> |
|[<span data-ttu-id="4ffd2-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ffd2-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ffd2-181">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ffd2-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="4ffd2-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-182">SourceProperty :String</span></span>

<span data-ttu-id="4ffd2-183">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="4ffd2-184">Type :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-184">Type:</span></span>

*   <span data-ttu-id="4ffd2-185">String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="4ffd2-186">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ffd2-186">Properties:</span></span>

|<span data-ttu-id="4ffd2-187">Nom</span><span class="sxs-lookup"><span data-stu-id="4ffd2-187">Name</span></span>| <span data-ttu-id="4ffd2-188">Type</span><span class="sxs-lookup"><span data-stu-id="4ffd2-188">Type</span></span>| <span data-ttu-id="4ffd2-189">Description</span><span class="sxs-lookup"><span data-stu-id="4ffd2-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="4ffd2-190">Chaîne</span><span class="sxs-lookup"><span data-stu-id="4ffd2-190">String</span></span>|<span data-ttu-id="4ffd2-191">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="4ffd2-192">String</span><span class="sxs-lookup"><span data-stu-id="4ffd2-192">String</span></span>|<span data-ttu-id="4ffd2-193">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="4ffd2-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ffd2-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ffd2-194">Requirements</span></span>

|<span data-ttu-id="4ffd2-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ffd2-195">Requirement</span></span>| <span data-ttu-id="4ffd2-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ffd2-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ffd2-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ffd2-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ffd2-198">1.0</span><span class="sxs-lookup"><span data-stu-id="4ffd2-198">1.0</span></span>|
|[<span data-ttu-id="4ffd2-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ffd2-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4ffd2-200">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ffd2-200">Compose or read</span></span>|