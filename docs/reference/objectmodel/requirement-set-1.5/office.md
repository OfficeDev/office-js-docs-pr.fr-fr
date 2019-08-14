---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: ''
ms.date: 08/13/2019
localization_priority: Normal
ms.openlocfilehash: 402737f0f6648e42f569906df59be0fa26991146
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395686"
---
# <a name="office"></a><span data-ttu-id="0f145-102">Office</span><span class="sxs-lookup"><span data-stu-id="0f145-102">Office</span></span>

<span data-ttu-id="0f145-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0f145-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0f145-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0f145-105">Requirements</span></span>

|<span data-ttu-id="0f145-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0f145-106">Requirement</span></span>| <span data-ttu-id="0f145-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="0f145-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f145-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0f145-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f145-109">1.0</span><span class="sxs-lookup"><span data-stu-id="0f145-109">1.0</span></span>|
|[<span data-ttu-id="0f145-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0f145-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f145-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0f145-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="0f145-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="0f145-112">Members and methods</span></span>

| <span data-ttu-id="0f145-113">Membre</span><span class="sxs-lookup"><span data-stu-id="0f145-113">Member</span></span> | <span data-ttu-id="0f145-114">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="0f145-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="0f145-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="0f145-116">Member</span><span class="sxs-lookup"><span data-stu-id="0f145-116">Member</span></span> |
| [<span data-ttu-id="0f145-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="0f145-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="0f145-118">Member</span><span class="sxs-lookup"><span data-stu-id="0f145-118">Member</span></span> |
| [<span data-ttu-id="0f145-119">EventType</span><span class="sxs-lookup"><span data-stu-id="0f145-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="0f145-120">Member</span><span class="sxs-lookup"><span data-stu-id="0f145-120">Member</span></span> |
| [<span data-ttu-id="0f145-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="0f145-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="0f145-122">Membre</span><span class="sxs-lookup"><span data-stu-id="0f145-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0f145-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="0f145-123">Namespaces</span></span>

<span data-ttu-id="0f145-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="0f145-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0f145-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): inclut un certain nombre d’énumérations, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="0f145-125">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5): Includes a number of enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

### <a name="members"></a><span data-ttu-id="0f145-126">Members</span><span class="sxs-lookup"><span data-stu-id="0f145-126">Members</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="0f145-127">AsyncResultStatus: chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-127">AsyncResultStatus: String</span></span>

<span data-ttu-id="0f145-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="0f145-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0f145-129">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-129">Type</span></span>

*   <span data-ttu-id="0f145-130">String</span><span class="sxs-lookup"><span data-stu-id="0f145-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f145-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0f145-131">Properties:</span></span>

|<span data-ttu-id="0f145-132">Nom</span><span class="sxs-lookup"><span data-stu-id="0f145-132">Name</span></span>| <span data-ttu-id="0f145-133">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-133">Type</span></span>| <span data-ttu-id="0f145-134">Description</span><span class="sxs-lookup"><span data-stu-id="0f145-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0f145-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-135">String</span></span>|<span data-ttu-id="0f145-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="0f145-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0f145-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-137">String</span></span>|<span data-ttu-id="0f145-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="0f145-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f145-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0f145-139">Requirements</span></span>

|<span data-ttu-id="0f145-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0f145-140">Requirement</span></span>| <span data-ttu-id="0f145-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="0f145-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f145-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0f145-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f145-143">1.0</span><span class="sxs-lookup"><span data-stu-id="0f145-143">1.0</span></span>|
|[<span data-ttu-id="0f145-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0f145-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f145-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0f145-145">Compose or Read</span></span>|

---

#### <a name="coerciontype-string"></a><span data-ttu-id="0f145-146">CoercionType: chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-146">CoercionType: String</span></span>

<span data-ttu-id="0f145-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="0f145-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0f145-148">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-148">Type</span></span>

*   <span data-ttu-id="0f145-149">String</span><span class="sxs-lookup"><span data-stu-id="0f145-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f145-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0f145-150">Properties:</span></span>

|<span data-ttu-id="0f145-151">Nom</span><span class="sxs-lookup"><span data-stu-id="0f145-151">Name</span></span>| <span data-ttu-id="0f145-152">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-152">Type</span></span>| <span data-ttu-id="0f145-153">Description</span><span class="sxs-lookup"><span data-stu-id="0f145-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0f145-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-154">String</span></span>|<span data-ttu-id="0f145-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="0f145-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0f145-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-156">String</span></span>|<span data-ttu-id="0f145-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="0f145-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f145-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0f145-158">Requirements</span></span>

|<span data-ttu-id="0f145-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0f145-159">Requirement</span></span>| <span data-ttu-id="0f145-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="0f145-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f145-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0f145-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f145-162">1.0</span><span class="sxs-lookup"><span data-stu-id="0f145-162">1.0</span></span>|
|[<span data-ttu-id="0f145-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0f145-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f145-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0f145-164">Compose or Read</span></span>|

---

#### <a name="eventtype-string"></a><span data-ttu-id="0f145-165">EventType: chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-165">EventType: String</span></span>

<span data-ttu-id="0f145-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="0f145-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="0f145-167">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-167">Type</span></span>

*   <span data-ttu-id="0f145-168">String</span><span class="sxs-lookup"><span data-stu-id="0f145-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f145-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0f145-169">Properties:</span></span>

| <span data-ttu-id="0f145-170">Nom</span><span class="sxs-lookup"><span data-stu-id="0f145-170">Name</span></span> | <span data-ttu-id="0f145-171">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-171">Type</span></span> | <span data-ttu-id="0f145-172">Description</span><span class="sxs-lookup"><span data-stu-id="0f145-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="0f145-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-173">String</span></span> | <span data-ttu-id="0f145-174">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="0f145-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="0f145-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0f145-175">Requirements</span></span>

|<span data-ttu-id="0f145-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0f145-176">Requirement</span></span>| <span data-ttu-id="0f145-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="0f145-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f145-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0f145-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f145-179">1,5</span><span class="sxs-lookup"><span data-stu-id="0f145-179">1.5</span></span> |
|[<span data-ttu-id="0f145-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0f145-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f145-181">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0f145-181">Compose or Read</span></span> |

---

#### <a name="sourceproperty-string"></a><span data-ttu-id="0f145-182">SourceProperty: chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-182">SourceProperty: String</span></span>

<span data-ttu-id="0f145-183">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="0f145-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0f145-184">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-184">Type</span></span>

*   <span data-ttu-id="0f145-185">String</span><span class="sxs-lookup"><span data-stu-id="0f145-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0f145-186">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="0f145-186">Properties:</span></span>

|<span data-ttu-id="0f145-187">Nom</span><span class="sxs-lookup"><span data-stu-id="0f145-187">Name</span></span>| <span data-ttu-id="0f145-188">Type</span><span class="sxs-lookup"><span data-stu-id="0f145-188">Type</span></span>| <span data-ttu-id="0f145-189">Description</span><span class="sxs-lookup"><span data-stu-id="0f145-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0f145-190">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-190">String</span></span>|<span data-ttu-id="0f145-191">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="0f145-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0f145-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="0f145-192">String</span></span>|<span data-ttu-id="0f145-193">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="0f145-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0f145-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="0f145-194">Requirements</span></span>

|<span data-ttu-id="0f145-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="0f145-195">Requirement</span></span>| <span data-ttu-id="0f145-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="0f145-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="0f145-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="0f145-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0f145-198">1.0</span><span class="sxs-lookup"><span data-stu-id="0f145-198">1.0</span></span>|
|[<span data-ttu-id="0f145-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="0f145-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0f145-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="0f145-200">Compose or Read</span></span>|
