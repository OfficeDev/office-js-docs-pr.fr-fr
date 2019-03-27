---
title: Espace de noms Office-ensemble de conditions requises 1,6
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: dde96f48863459da5072d6b4864169f198264133
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870806"
---
# <a name="office"></a><span data-ttu-id="8a1f6-102">Office</span><span class="sxs-lookup"><span data-stu-id="8a1f6-102">Office</span></span>

<span data-ttu-id="8a1f6-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="8a1f6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a1f6-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8a1f6-105">Requirements</span></span>

|<span data-ttu-id="8a1f6-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8a1f6-106">Requirement</span></span>| <span data-ttu-id="8a1f6-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="8a1f6-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a1f6-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8a1f6-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a1f6-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8a1f6-109">1.0</span></span>|
|[<span data-ttu-id="8a1f6-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8a1f6-110">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a1f6-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8a1f6-111">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8a1f6-112">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="8a1f6-112">Members and methods</span></span>

| <span data-ttu-id="8a1f6-113">Membre</span><span class="sxs-lookup"><span data-stu-id="8a1f6-113">Member</span></span> | <span data-ttu-id="8a1f6-114">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-114">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8a1f6-115">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8a1f6-115">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8a1f6-116">Member</span><span class="sxs-lookup"><span data-stu-id="8a1f6-116">Member</span></span> |
| [<span data-ttu-id="8a1f6-117">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8a1f6-117">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8a1f6-118">Member</span><span class="sxs-lookup"><span data-stu-id="8a1f6-118">Member</span></span> |
| [<span data-ttu-id="8a1f6-119">EventType</span><span class="sxs-lookup"><span data-stu-id="8a1f6-119">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8a1f6-120">Member</span><span class="sxs-lookup"><span data-stu-id="8a1f6-120">Member</span></span> |
| [<span data-ttu-id="8a1f6-121">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8a1f6-121">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8a1f6-122">Membre</span><span class="sxs-lookup"><span data-stu-id="8a1f6-122">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8a1f6-123">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="8a1f6-123">Namespaces</span></span>

<span data-ttu-id="8a1f6-124">[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-124">[context](office.context.md): Provides shared interfaces from the Office Add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="8a1f6-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-125">[MailboxEnums](/javascript/api/outlook_1_6/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="8a1f6-126">Membres</span><span class="sxs-lookup"><span data-stu-id="8a1f6-126">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="8a1f6-127">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-127">AsyncResultStatus :String</span></span>

<span data-ttu-id="8a1f6-128">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-128">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8a1f6-129">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-129">Type</span></span>

*   <span data-ttu-id="8a1f6-130">String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-130">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8a1f6-131">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8a1f6-131">Properties:</span></span>

|<span data-ttu-id="8a1f6-132">Nom</span><span class="sxs-lookup"><span data-stu-id="8a1f6-132">Name</span></span>| <span data-ttu-id="8a1f6-133">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-133">Type</span></span>| <span data-ttu-id="8a1f6-134">Description</span><span class="sxs-lookup"><span data-stu-id="8a1f6-134">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8a1f6-135">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8a1f6-135">String</span></span>|<span data-ttu-id="8a1f6-136">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-136">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8a1f6-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8a1f6-137">String</span></span>|<span data-ttu-id="8a1f6-138">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-138">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a1f6-139">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8a1f6-139">Requirements</span></span>

|<span data-ttu-id="8a1f6-140">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8a1f6-140">Requirement</span></span>| <span data-ttu-id="8a1f6-141">Valeur</span><span class="sxs-lookup"><span data-stu-id="8a1f6-141">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a1f6-142">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8a1f6-142">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a1f6-143">1.0</span><span class="sxs-lookup"><span data-stu-id="8a1f6-143">1.0</span></span>|
|[<span data-ttu-id="8a1f6-144">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8a1f6-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a1f6-145">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8a1f6-145">Compose or Read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="8a1f6-146">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-146">CoercionType :String</span></span>

<span data-ttu-id="8a1f6-147">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-147">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8a1f6-148">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-148">Type</span></span>

*   <span data-ttu-id="8a1f6-149">String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-149">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8a1f6-150">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8a1f6-150">Properties:</span></span>

|<span data-ttu-id="8a1f6-151">Nom</span><span class="sxs-lookup"><span data-stu-id="8a1f6-151">Name</span></span>| <span data-ttu-id="8a1f6-152">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-152">Type</span></span>| <span data-ttu-id="8a1f6-153">Description</span><span class="sxs-lookup"><span data-stu-id="8a1f6-153">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8a1f6-154">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8a1f6-154">String</span></span>|<span data-ttu-id="8a1f6-155">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-155">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8a1f6-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8a1f6-156">String</span></span>|<span data-ttu-id="8a1f6-157">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-157">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a1f6-158">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8a1f6-158">Requirements</span></span>

|<span data-ttu-id="8a1f6-159">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8a1f6-159">Requirement</span></span>| <span data-ttu-id="8a1f6-160">Valeur</span><span class="sxs-lookup"><span data-stu-id="8a1f6-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a1f6-161">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8a1f6-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a1f6-162">1.0</span><span class="sxs-lookup"><span data-stu-id="8a1f6-162">1.0</span></span>|
|[<span data-ttu-id="8a1f6-163">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8a1f6-163">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a1f6-164">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8a1f6-164">Compose or Read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="8a1f6-165">EventType :String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-165">EventType :String</span></span>

<span data-ttu-id="8a1f6-166">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-166">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8a1f6-167">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-167">Type</span></span>

*   <span data-ttu-id="8a1f6-168">String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-168">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8a1f6-169">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8a1f6-169">Properties:</span></span>

| <span data-ttu-id="8a1f6-170">Nom</span><span class="sxs-lookup"><span data-stu-id="8a1f6-170">Name</span></span> | <span data-ttu-id="8a1f6-171">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-171">Type</span></span> | <span data-ttu-id="8a1f6-172">Description</span><span class="sxs-lookup"><span data-stu-id="8a1f6-172">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="8a1f6-173">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8a1f6-173">String</span></span> | <span data-ttu-id="8a1f6-174">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-174">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8a1f6-175">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8a1f6-175">Requirements</span></span>

|<span data-ttu-id="8a1f6-176">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8a1f6-176">Requirement</span></span>| <span data-ttu-id="8a1f6-177">Valeur</span><span class="sxs-lookup"><span data-stu-id="8a1f6-177">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a1f6-178">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8a1f6-178">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a1f6-179">1,5</span><span class="sxs-lookup"><span data-stu-id="8a1f6-179">1.5</span></span> |
|[<span data-ttu-id="8a1f6-180">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8a1f6-180">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a1f6-181">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8a1f6-181">Compose or Read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="8a1f6-182">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-182">SourceProperty :String</span></span>

<span data-ttu-id="8a1f6-183">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-183">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8a1f6-184">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-184">Type</span></span>

*   <span data-ttu-id="8a1f6-185">String</span><span class="sxs-lookup"><span data-stu-id="8a1f6-185">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8a1f6-186">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="8a1f6-186">Properties:</span></span>

|<span data-ttu-id="8a1f6-187">Nom</span><span class="sxs-lookup"><span data-stu-id="8a1f6-187">Name</span></span>| <span data-ttu-id="8a1f6-188">Type</span><span class="sxs-lookup"><span data-stu-id="8a1f6-188">Type</span></span>| <span data-ttu-id="8a1f6-189">Description</span><span class="sxs-lookup"><span data-stu-id="8a1f6-189">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8a1f6-190">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8a1f6-190">String</span></span>|<span data-ttu-id="8a1f6-191">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-191">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8a1f6-192">Chaîne</span><span class="sxs-lookup"><span data-stu-id="8a1f6-192">String</span></span>|<span data-ttu-id="8a1f6-193">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="8a1f6-193">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8a1f6-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="8a1f6-194">Requirements</span></span>

|<span data-ttu-id="8a1f6-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="8a1f6-195">Requirement</span></span>| <span data-ttu-id="8a1f6-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="8a1f6-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a1f6-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="8a1f6-197">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a1f6-198">1.0</span><span class="sxs-lookup"><span data-stu-id="8a1f6-198">1.0</span></span>|
|[<span data-ttu-id="8a1f6-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="8a1f6-199">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a1f6-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="8a1f6-200">Compose or Read</span></span>|
