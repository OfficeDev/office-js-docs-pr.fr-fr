---
title: Espace de noms Office-ensemble de conditions requises 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: c9f769550ad2c4994545e51d140b6ea6e67761bc
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067936"
---
# <a name="office"></a><span data-ttu-id="19d76-102">Office</span><span class="sxs-lookup"><span data-stu-id="19d76-102">Office</span></span>

<span data-ttu-id="19d76-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="19d76-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="19d76-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="19d76-105">Requirements</span></span>

|<span data-ttu-id="19d76-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="19d76-106">Requirement</span></span>| <span data-ttu-id="19d76-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="19d76-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="19d76-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="19d76-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19d76-109">1.0</span><span class="sxs-lookup"><span data-stu-id="19d76-109">1.0</span></span>|
|[<span data-ttu-id="19d76-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="19d76-110">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19d76-111">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="19d76-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="19d76-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="19d76-112">Members and methods</span></span>

| <span data-ttu-id="19d76-113">Membre</span><span class="sxs-lookup"><span data-stu-id="19d76-113">Member</span></span> | <span data-ttu-id="19d76-114">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="19d76-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="19d76-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="19d76-116">Membre</span><span class="sxs-lookup"><span data-stu-id="19d76-116">Member</span></span> |
| [<span data-ttu-id="19d76-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="19d76-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="19d76-118">Membre</span><span class="sxs-lookup"><span data-stu-id="19d76-118">Member</span></span> |
| [<span data-ttu-id="19d76-119">EventType</span><span class="sxs-lookup"><span data-stu-id="19d76-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="19d76-120">Membre</span><span class="sxs-lookup"><span data-stu-id="19d76-120">Member</span></span> |
| [<span data-ttu-id="19d76-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="19d76-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="19d76-122">Membre</span><span class="sxs-lookup"><span data-stu-id="19d76-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="19d76-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="19d76-123">Namespaces</span></span>

<span data-ttu-id="19d76-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="19d76-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="19d76-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="19d76-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="19d76-126">Membres</span><span class="sxs-lookup"><span data-stu-id="19d76-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="19d76-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="19d76-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="19d76-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="19d76-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="19d76-129">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-129">Type</span></span>

*   <span data-ttu-id="19d76-130">String</span><span class="sxs-lookup"><span data-stu-id="19d76-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19d76-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="19d76-131">Properties:</span></span>

|<span data-ttu-id="19d76-132">Nom</span><span class="sxs-lookup"><span data-stu-id="19d76-132">Name</span></span>| <span data-ttu-id="19d76-133">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-133">Type</span></span>| <span data-ttu-id="19d76-134">Description</span><span class="sxs-lookup"><span data-stu-id="19d76-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="19d76-135">String</span><span class="sxs-lookup"><span data-stu-id="19d76-135">String</span></span>|<span data-ttu-id="19d76-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="19d76-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="19d76-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="19d76-137">String</span></span>|<span data-ttu-id="19d76-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="19d76-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19d76-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="19d76-139">Requirements</span></span>

|<span data-ttu-id="19d76-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="19d76-140">Requirement</span></span>| <span data-ttu-id="19d76-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="19d76-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="19d76-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="19d76-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19d76-143">1.0</span><span class="sxs-lookup"><span data-stu-id="19d76-143">1.0</span></span>|
|[<span data-ttu-id="19d76-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="19d76-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19d76-145">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="19d76-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="19d76-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="19d76-146">CoercionType :String</span></span>

<span data-ttu-id="19d76-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="19d76-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="19d76-148">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-148">Type</span></span>

*   <span data-ttu-id="19d76-149">String</span><span class="sxs-lookup"><span data-stu-id="19d76-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19d76-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="19d76-150">Properties:</span></span>

|<span data-ttu-id="19d76-151">Nom</span><span class="sxs-lookup"><span data-stu-id="19d76-151">Name</span></span>| <span data-ttu-id="19d76-152">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-152">Type</span></span>| <span data-ttu-id="19d76-153">Description</span><span class="sxs-lookup"><span data-stu-id="19d76-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="19d76-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="19d76-154">String</span></span>|<span data-ttu-id="19d76-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="19d76-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="19d76-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="19d76-156">String</span></span>|<span data-ttu-id="19d76-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="19d76-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19d76-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="19d76-158">Requirements</span></span>

|<span data-ttu-id="19d76-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="19d76-159">Requirement</span></span>| <span data-ttu-id="19d76-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="19d76-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="19d76-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="19d76-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19d76-162">1.0</span><span class="sxs-lookup"><span data-stu-id="19d76-162">1.0</span></span>|
|[<span data-ttu-id="19d76-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="19d76-163">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19d76-164">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="19d76-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="19d76-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="19d76-165">EventType :String</span></span>

<span data-ttu-id="19d76-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="19d76-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="19d76-167">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-167">Type</span></span>

*   <span data-ttu-id="19d76-168">String</span><span class="sxs-lookup"><span data-stu-id="19d76-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19d76-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="19d76-169">Properties:</span></span>

| <span data-ttu-id="19d76-170">Nom</span><span class="sxs-lookup"><span data-stu-id="19d76-170">Name</span></span> | <span data-ttu-id="19d76-171">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-171">Type</span></span> | <span data-ttu-id="19d76-172">Description</span><span class="sxs-lookup"><span data-stu-id="19d76-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="19d76-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="19d76-173">String</span></span> | <span data-ttu-id="19d76-174">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="19d76-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="19d76-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="19d76-175">Requirements</span></span>

|<span data-ttu-id="19d76-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="19d76-176">Requirement</span></span>| <span data-ttu-id="19d76-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="19d76-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="19d76-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="19d76-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19d76-179">1,5</span><span class="sxs-lookup"><span data-stu-id="19d76-179">1.5</span></span> |
|[<span data-ttu-id="19d76-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="19d76-180">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19d76-181">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="19d76-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="19d76-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="19d76-182">SourceProperty :String</span></span>

<span data-ttu-id="19d76-183">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="19d76-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="19d76-184">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-184">Type</span></span>

*   <span data-ttu-id="19d76-185">String</span><span class="sxs-lookup"><span data-stu-id="19d76-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="19d76-186">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="19d76-186">Properties:</span></span>

|<span data-ttu-id="19d76-187">Nom</span><span class="sxs-lookup"><span data-stu-id="19d76-187">Name</span></span>| <span data-ttu-id="19d76-188">Type</span><span class="sxs-lookup"><span data-stu-id="19d76-188">Type</span></span>| <span data-ttu-id="19d76-189">Description</span><span class="sxs-lookup"><span data-stu-id="19d76-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="19d76-190">Chaîne</span><span class="sxs-lookup"><span data-stu-id="19d76-190">String</span></span>|<span data-ttu-id="19d76-191">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="19d76-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="19d76-192">String</span><span class="sxs-lookup"><span data-stu-id="19d76-192">String</span></span>|<span data-ttu-id="19d76-193">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="19d76-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19d76-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="19d76-194">Requirements</span></span>

|<span data-ttu-id="19d76-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="19d76-195">Requirement</span></span>| <span data-ttu-id="19d76-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="19d76-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="19d76-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="19d76-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19d76-198">1.0</span><span class="sxs-lookup"><span data-stu-id="19d76-198">1.0</span></span>|
|[<span data-ttu-id="19d76-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="19d76-199">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="19d76-200">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="19d76-200">Compose or Read</span></span>|
