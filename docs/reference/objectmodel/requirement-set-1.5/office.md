---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: d8c51646818681629fa0c184962776beffe22a55
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/24/2019
ms.locfileid: "32450260"
---
# <a name="office"></a><span data-ttu-id="c5e1c-102">Office</span><span class="sxs-lookup"><span data-stu-id="c5e1c-102">Office</span></span>

<span data-ttu-id="c5e1c-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c5e1c-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c5e1c-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c5e1c-105">Requirements</span></span>

|<span data-ttu-id="c5e1c-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c5e1c-106">Requirement</span></span>| <span data-ttu-id="c5e1c-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="c5e1c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5e1c-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c5e1c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5e1c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="c5e1c-109">1.0</span></span>|
|[<span data-ttu-id="c5e1c-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c5e1c-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5e1c-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c5e1c-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c5e1c-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c5e1c-112">Members and methods</span></span>

| <span data-ttu-id="c5e1c-113">Membre</span><span class="sxs-lookup"><span data-stu-id="c5e1c-113">Member</span></span> | <span data-ttu-id="c5e1c-114">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c5e1c-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c5e1c-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c5e1c-116">Member</span><span class="sxs-lookup"><span data-stu-id="c5e1c-116">Member</span></span> |
| [<span data-ttu-id="c5e1c-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c5e1c-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c5e1c-118">Member</span><span class="sxs-lookup"><span data-stu-id="c5e1c-118">Member</span></span> |
| [<span data-ttu-id="c5e1c-119">EventType</span><span class="sxs-lookup"><span data-stu-id="c5e1c-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c5e1c-120">Member</span><span class="sxs-lookup"><span data-stu-id="c5e1c-120">Member</span></span> |
| [<span data-ttu-id="c5e1c-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c5e1c-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c5e1c-122">Membre</span><span class="sxs-lookup"><span data-stu-id="c5e1c-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c5e1c-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c5e1c-123">Namespaces</span></span>

<span data-ttu-id="c5e1c-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c5e1c-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-125">[MailboxEnums](/javascript/api/outlook_1_5/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c5e1c-126">Membres</span><span class="sxs-lookup"><span data-stu-id="c5e1c-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c5e1c-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="c5e1c-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c5e1c-129">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-129">Type</span></span>

*   <span data-ttu-id="c5e1c-130">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c5e1c-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c5e1c-131">Properties:</span></span>

|<span data-ttu-id="c5e1c-132">Nom</span><span class="sxs-lookup"><span data-stu-id="c5e1c-132">Name</span></span>| <span data-ttu-id="c5e1c-133">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-133">Type</span></span>| <span data-ttu-id="c5e1c-134">Description</span><span class="sxs-lookup"><span data-stu-id="c5e1c-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c5e1c-135">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-135">String</span></span>|<span data-ttu-id="c5e1c-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c5e1c-137">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-137">String</span></span>|<span data-ttu-id="c5e1c-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c5e1c-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c5e1c-139">Requirements</span></span>

|<span data-ttu-id="c5e1c-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c5e1c-140">Requirement</span></span>| <span data-ttu-id="c5e1c-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="c5e1c-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5e1c-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c5e1c-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5e1c-143">1.0</span><span class="sxs-lookup"><span data-stu-id="c5e1c-143">1.0</span></span>|
|[<span data-ttu-id="c5e1c-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c5e1c-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5e1c-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c5e1c-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="c5e1c-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-146">CoercionType :String</span></span>

<span data-ttu-id="c5e1c-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c5e1c-148">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-148">Type</span></span>

*   <span data-ttu-id="c5e1c-149">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c5e1c-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c5e1c-150">Properties:</span></span>

|<span data-ttu-id="c5e1c-151">Nom</span><span class="sxs-lookup"><span data-stu-id="c5e1c-151">Name</span></span>| <span data-ttu-id="c5e1c-152">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-152">Type</span></span>| <span data-ttu-id="c5e1c-153">Description</span><span class="sxs-lookup"><span data-stu-id="c5e1c-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c5e1c-154">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-154">String</span></span>|<span data-ttu-id="c5e1c-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c5e1c-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="c5e1c-156">String</span></span>|<span data-ttu-id="c5e1c-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c5e1c-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c5e1c-158">Requirements</span></span>

|<span data-ttu-id="c5e1c-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c5e1c-159">Requirement</span></span>| <span data-ttu-id="c5e1c-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="c5e1c-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5e1c-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c5e1c-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5e1c-162">1.0</span><span class="sxs-lookup"><span data-stu-id="c5e1c-162">1.0</span></span>|
|[<span data-ttu-id="c5e1c-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c5e1c-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5e1c-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c5e1c-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="c5e1c-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-165">EventType :String</span></span>

<span data-ttu-id="c5e1c-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c5e1c-167">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-167">Type</span></span>

*   <span data-ttu-id="c5e1c-168">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c5e1c-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c5e1c-169">Properties:</span></span>

| <span data-ttu-id="c5e1c-170">Nom</span><span class="sxs-lookup"><span data-stu-id="c5e1c-170">Name</span></span> | <span data-ttu-id="c5e1c-171">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-171">Type</span></span> | <span data-ttu-id="c5e1c-172">Description</span><span class="sxs-lookup"><span data-stu-id="c5e1c-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="c5e1c-173">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-173">String</span></span> | <span data-ttu-id="c5e1c-174">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c5e1c-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c5e1c-175">Requirements</span></span>

|<span data-ttu-id="c5e1c-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c5e1c-176">Requirement</span></span>| <span data-ttu-id="c5e1c-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="c5e1c-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5e1c-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c5e1c-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5e1c-179">1,5</span><span class="sxs-lookup"><span data-stu-id="c5e1c-179">1.5</span></span> |
|[<span data-ttu-id="c5e1c-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c5e1c-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5e1c-181">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c5e1c-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="c5e1c-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-182">SourceProperty :String</span></span>

<span data-ttu-id="c5e1c-183">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c5e1c-184">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-184">Type</span></span>

*   <span data-ttu-id="c5e1c-185">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c5e1c-186">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c5e1c-186">Properties:</span></span>

|<span data-ttu-id="c5e1c-187">Nom</span><span class="sxs-lookup"><span data-stu-id="c5e1c-187">Name</span></span>| <span data-ttu-id="c5e1c-188">Type</span><span class="sxs-lookup"><span data-stu-id="c5e1c-188">Type</span></span>| <span data-ttu-id="c5e1c-189">Description</span><span class="sxs-lookup"><span data-stu-id="c5e1c-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c5e1c-190">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-190">String</span></span>|<span data-ttu-id="c5e1c-191">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c5e1c-192">String</span><span class="sxs-lookup"><span data-stu-id="c5e1c-192">String</span></span>|<span data-ttu-id="c5e1c-193">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="c5e1c-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c5e1c-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c5e1c-194">Requirements</span></span>

|<span data-ttu-id="c5e1c-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c5e1c-195">Requirement</span></span>| <span data-ttu-id="c5e1c-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="c5e1c-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="c5e1c-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c5e1c-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c5e1c-198">1.0</span><span class="sxs-lookup"><span data-stu-id="c5e1c-198">1.0</span></span>|
|[<span data-ttu-id="c5e1c-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c5e1c-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c5e1c-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c5e1c-200">Compose or Read</span></span>|
