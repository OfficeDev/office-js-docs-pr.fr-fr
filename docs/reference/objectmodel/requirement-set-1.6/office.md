---
title: Espace de noms Office-ensemble de conditions requises 1.6
description: ''
ms.date: 11/08/2018
ms.openlocfilehash: 73411efee9dcfffa5f9f0fa9de85dafc31a4173a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432283"
---
# <a name="office"></a><span data-ttu-id="c1621-102">Office</span><span class="sxs-lookup"><span data-stu-id="c1621-102">Office</span></span>

<span data-ttu-id="c1621-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API partagée](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c1621-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1621-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1621-105">Requirements</span></span>

|<span data-ttu-id="c1621-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1621-106">Requirement</span></span>| <span data-ttu-id="c1621-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1621-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1621-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1621-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1621-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c1621-109">1.0</span></span>|
|[<span data-ttu-id="c1621-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1621-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1621-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1621-111">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c1621-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c1621-112">Members and methods</span></span>

| <span data-ttu-id="c1621-113">Membre</span><span class="sxs-lookup"><span data-stu-id="c1621-113">Member</span></span> | <span data-ttu-id="c1621-114">Type</span><span class="sxs-lookup"><span data-stu-id="c1621-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c1621-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c1621-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c1621-116">Membre</span><span class="sxs-lookup"><span data-stu-id="c1621-116">Member</span></span> |
| [<span data-ttu-id="c1621-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c1621-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c1621-118">Membre</span><span class="sxs-lookup"><span data-stu-id="c1621-118">Member</span></span> |
| [<span data-ttu-id="c1621-119">EventType</span><span class="sxs-lookup"><span data-stu-id="c1621-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c1621-120">Membre</span><span class="sxs-lookup"><span data-stu-id="c1621-120">Member</span></span> |
| [<span data-ttu-id="c1621-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c1621-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c1621-122">Membre</span><span class="sxs-lookup"><span data-stu-id="c1621-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c1621-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c1621-123">Namespaces</span></span>

<span data-ttu-id="c1621-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="c1621-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c1621-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c1621-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c1621-126">Membres</span><span class="sxs-lookup"><span data-stu-id="c1621-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c1621-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c1621-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="c1621-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c1621-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c1621-129">Type :</span><span class="sxs-lookup"><span data-stu-id="c1621-129">Type:</span></span>

*   <span data-ttu-id="c1621-130">String</span><span class="sxs-lookup"><span data-stu-id="c1621-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1621-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c1621-131">Properties:</span></span>

|<span data-ttu-id="c1621-132">Nom</span><span class="sxs-lookup"><span data-stu-id="c1621-132">Name</span></span>| <span data-ttu-id="c1621-133">Type</span><span class="sxs-lookup"><span data-stu-id="c1621-133">Type</span></span>| <span data-ttu-id="c1621-134">Description</span><span class="sxs-lookup"><span data-stu-id="c1621-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c1621-135">String</span><span class="sxs-lookup"><span data-stu-id="c1621-135">String</span></span>|<span data-ttu-id="c1621-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="c1621-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c1621-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1621-137">String</span></span>|<span data-ttu-id="c1621-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="c1621-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1621-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1621-139">Requirements</span></span>

|<span data-ttu-id="c1621-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1621-140">Requirement</span></span>| <span data-ttu-id="c1621-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1621-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1621-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1621-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1621-143">1.0</span><span class="sxs-lookup"><span data-stu-id="c1621-143">1.0</span></span>|
|[<span data-ttu-id="c1621-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1621-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1621-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1621-145">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="c1621-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c1621-146">CoercionType :String</span></span>

<span data-ttu-id="c1621-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c1621-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c1621-148">Type :</span><span class="sxs-lookup"><span data-stu-id="c1621-148">Type:</span></span>

*   <span data-ttu-id="c1621-149">String</span><span class="sxs-lookup"><span data-stu-id="c1621-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1621-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c1621-150">Properties:</span></span>

|<span data-ttu-id="c1621-151">Nom</span><span class="sxs-lookup"><span data-stu-id="c1621-151">Name</span></span>| <span data-ttu-id="c1621-152">Type</span><span class="sxs-lookup"><span data-stu-id="c1621-152">Type</span></span>| <span data-ttu-id="c1621-153">Description</span><span class="sxs-lookup"><span data-stu-id="c1621-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c1621-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1621-154">String</span></span>|<span data-ttu-id="c1621-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="c1621-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c1621-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1621-156">String</span></span>|<span data-ttu-id="c1621-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="c1621-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1621-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1621-158">Requirements</span></span>

|<span data-ttu-id="c1621-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1621-159">Requirement</span></span>| <span data-ttu-id="c1621-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1621-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1621-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1621-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1621-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c1621-162">1.0</span></span>|
|[<span data-ttu-id="c1621-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1621-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1621-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1621-164">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="c1621-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="c1621-165">EventType :String</span></span>

<span data-ttu-id="c1621-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="c1621-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c1621-167">Type :</span><span class="sxs-lookup"><span data-stu-id="c1621-167">Type:</span></span>

*   <span data-ttu-id="c1621-168">String</span><span class="sxs-lookup"><span data-stu-id="c1621-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1621-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c1621-169">Properties:</span></span>

| <span data-ttu-id="c1621-170">Nom</span><span class="sxs-lookup"><span data-stu-id="c1621-170">Name</span></span> | <span data-ttu-id="c1621-171">Type</span><span class="sxs-lookup"><span data-stu-id="c1621-171">Type</span></span> | <span data-ttu-id="c1621-172">Description</span><span class="sxs-lookup"><span data-stu-id="c1621-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="c1621-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1621-173">String</span></span> | <span data-ttu-id="c1621-174">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="c1621-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c1621-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1621-175">Requirements</span></span>

|<span data-ttu-id="c1621-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1621-176">Requirement</span></span>| <span data-ttu-id="c1621-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1621-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1621-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1621-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1621-179">1,5</span><span class="sxs-lookup"><span data-stu-id="c1621-179">1.5</span></span> |
|[<span data-ttu-id="c1621-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1621-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1621-181">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1621-181">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="c1621-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c1621-182">SourceProperty :String</span></span>

<span data-ttu-id="c1621-183">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c1621-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c1621-184">Type :</span><span class="sxs-lookup"><span data-stu-id="c1621-184">Type:</span></span>

*   <span data-ttu-id="c1621-185">String</span><span class="sxs-lookup"><span data-stu-id="c1621-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c1621-186">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c1621-186">Properties:</span></span>

|<span data-ttu-id="c1621-187">Nom</span><span class="sxs-lookup"><span data-stu-id="c1621-187">Name</span></span>| <span data-ttu-id="c1621-188">Type</span><span class="sxs-lookup"><span data-stu-id="c1621-188">Type</span></span>| <span data-ttu-id="c1621-189">Description</span><span class="sxs-lookup"><span data-stu-id="c1621-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c1621-190">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c1621-190">String</span></span>|<span data-ttu-id="c1621-191">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="c1621-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c1621-192">String</span><span class="sxs-lookup"><span data-stu-id="c1621-192">String</span></span>|<span data-ttu-id="c1621-193">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="c1621-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c1621-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c1621-194">Requirements</span></span>

|<span data-ttu-id="c1621-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c1621-195">Requirement</span></span>| <span data-ttu-id="c1621-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="c1621-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1621-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c1621-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1621-198">1.0</span><span class="sxs-lookup"><span data-stu-id="c1621-198">1.0</span></span>|
|[<span data-ttu-id="c1621-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c1621-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1621-200">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="c1621-200">Compose or read</span></span>|