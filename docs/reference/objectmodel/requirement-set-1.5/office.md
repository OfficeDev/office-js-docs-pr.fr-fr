---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 7cc8e6acc60c28b44ec7a2b91bb5e388b2618a31
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554723"
---
# <a name="office"></a><span data-ttu-id="1a795-102">Office</span><span class="sxs-lookup"><span data-stu-id="1a795-102">Office</span></span>

<span data-ttu-id="1a795-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="1a795-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1a795-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1a795-105">Requirements</span></span>

|<span data-ttu-id="1a795-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a795-106">Requirement</span></span>| <span data-ttu-id="1a795-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="1a795-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a795-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1a795-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1a795-109">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-109">1.1</span></span>|
|[<span data-ttu-id="1a795-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1a795-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1a795-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1a795-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="1a795-112">Properties</span></span>

| <span data-ttu-id="1a795-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="1a795-113">Property</span></span> | <span data-ttu-id="1a795-114">Modes</span><span class="sxs-lookup"><span data-stu-id="1a795-114">Modes</span></span> | <span data-ttu-id="1a795-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="1a795-115">Return type</span></span> | <span data-ttu-id="1a795-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="1a795-116">Minimum</span></span><br><span data-ttu-id="1a795-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a795-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1a795-118">context</span><span class="sxs-lookup"><span data-stu-id="1a795-118">context</span></span>](office.context.md) | <span data-ttu-id="1a795-119">Composition</span><span class="sxs-lookup"><span data-stu-id="1a795-119">Compose</span></span><br><span data-ttu-id="1a795-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-120">Read</span></span> | [<span data-ttu-id="1a795-121">Context</span><span class="sxs-lookup"><span data-stu-id="1a795-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5) | [<span data-ttu-id="1a795-122">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="1a795-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="1a795-123">Enumerations</span></span>

| <span data-ttu-id="1a795-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="1a795-124">Enumeration</span></span> | <span data-ttu-id="1a795-125">Modes</span><span class="sxs-lookup"><span data-stu-id="1a795-125">Modes</span></span> | <span data-ttu-id="1a795-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="1a795-126">Return type</span></span> | <span data-ttu-id="1a795-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="1a795-127">Minimum</span></span><br><span data-ttu-id="1a795-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a795-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1a795-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1a795-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1a795-130">Composition</span><span class="sxs-lookup"><span data-stu-id="1a795-130">Compose</span></span><br><span data-ttu-id="1a795-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-131">Read</span></span> | <span data-ttu-id="1a795-132">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-132">String</span></span> | [<span data-ttu-id="1a795-133">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1a795-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1a795-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1a795-135">Composition</span><span class="sxs-lookup"><span data-stu-id="1a795-135">Compose</span></span><br><span data-ttu-id="1a795-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-136">Read</span></span> | <span data-ttu-id="1a795-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-137">String</span></span> | [<span data-ttu-id="1a795-138">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1a795-139">EventType</span><span class="sxs-lookup"><span data-stu-id="1a795-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="1a795-140">Composition</span><span class="sxs-lookup"><span data-stu-id="1a795-140">Compose</span></span><br><span data-ttu-id="1a795-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-141">Read</span></span> | <span data-ttu-id="1a795-142">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-142">String</span></span> | [<span data-ttu-id="1a795-143">1,5</span><span class="sxs-lookup"><span data-stu-id="1a795-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1a795-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1a795-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1a795-145">Composition</span><span class="sxs-lookup"><span data-stu-id="1a795-145">Compose</span></span><br><span data-ttu-id="1a795-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-146">Read</span></span> | <span data-ttu-id="1a795-147">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-147">String</span></span> | [<span data-ttu-id="1a795-148">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="1a795-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="1a795-149">Namespaces</span></span>

<span data-ttu-id="1a795-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="1a795-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="1a795-151">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="1a795-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="1a795-152">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="1a795-153">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="1a795-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1a795-154">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-154">Type</span></span>

*   <span data-ttu-id="1a795-155">String</span><span class="sxs-lookup"><span data-stu-id="1a795-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1a795-156">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="1a795-156">Properties:</span></span>

|<span data-ttu-id="1a795-157">Nom</span><span class="sxs-lookup"><span data-stu-id="1a795-157">Name</span></span>| <span data-ttu-id="1a795-158">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-158">Type</span></span>| <span data-ttu-id="1a795-159">Description</span><span class="sxs-lookup"><span data-stu-id="1a795-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1a795-160">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-160">String</span></span>|<span data-ttu-id="1a795-161">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="1a795-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1a795-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-162">String</span></span>|<span data-ttu-id="1a795-163">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="1a795-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a795-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1a795-164">Requirements</span></span>

|<span data-ttu-id="1a795-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a795-165">Requirement</span></span>| <span data-ttu-id="1a795-166">Valeur</span><span class="sxs-lookup"><span data-stu-id="1a795-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a795-167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1a795-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1a795-168">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-168">1.1</span></span>|
|[<span data-ttu-id="1a795-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1a795-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1a795-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="1a795-171">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-171">CoercionType: String</span></span>

<span data-ttu-id="1a795-172">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="1a795-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1a795-173">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-173">Type</span></span>

*   <span data-ttu-id="1a795-174">String</span><span class="sxs-lookup"><span data-stu-id="1a795-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1a795-175">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="1a795-175">Properties:</span></span>

|<span data-ttu-id="1a795-176">Nom</span><span class="sxs-lookup"><span data-stu-id="1a795-176">Name</span></span>| <span data-ttu-id="1a795-177">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-177">Type</span></span>| <span data-ttu-id="1a795-178">Description</span><span class="sxs-lookup"><span data-stu-id="1a795-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1a795-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-179">String</span></span>|<span data-ttu-id="1a795-180">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="1a795-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1a795-181">String</span><span class="sxs-lookup"><span data-stu-id="1a795-181">String</span></span>|<span data-ttu-id="1a795-182">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="1a795-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a795-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1a795-183">Requirements</span></span>

|<span data-ttu-id="1a795-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a795-184">Requirement</span></span>| <span data-ttu-id="1a795-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="1a795-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a795-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1a795-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1a795-187">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-187">1.1</span></span>|
|[<span data-ttu-id="1a795-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1a795-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1a795-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="1a795-190">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-190">EventType: String</span></span>

<span data-ttu-id="1a795-191">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="1a795-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="1a795-192">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-192">Type</span></span>

*   <span data-ttu-id="1a795-193">String</span><span class="sxs-lookup"><span data-stu-id="1a795-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1a795-194">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="1a795-194">Properties:</span></span>

| <span data-ttu-id="1a795-195">Nom</span><span class="sxs-lookup"><span data-stu-id="1a795-195">Name</span></span> | <span data-ttu-id="1a795-196">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-196">Type</span></span> | <span data-ttu-id="1a795-197">Description</span><span class="sxs-lookup"><span data-stu-id="1a795-197">Description</span></span> | <span data-ttu-id="1a795-198">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="1a795-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="1a795-199">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-199">String</span></span> | <span data-ttu-id="1a795-200">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="1a795-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="1a795-201">1,5</span><span class="sxs-lookup"><span data-stu-id="1a795-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1a795-202">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1a795-202">Requirements</span></span>

|<span data-ttu-id="1a795-203">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a795-203">Requirement</span></span>| <span data-ttu-id="1a795-204">Valeur</span><span class="sxs-lookup"><span data-stu-id="1a795-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a795-205">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1a795-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1a795-206">1,5</span><span class="sxs-lookup"><span data-stu-id="1a795-206">1.5</span></span> |
|[<span data-ttu-id="1a795-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1a795-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1a795-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="1a795-209">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-209">SourceProperty: String</span></span>

<span data-ttu-id="1a795-210">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="1a795-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1a795-211">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-211">Type</span></span>

*   <span data-ttu-id="1a795-212">String</span><span class="sxs-lookup"><span data-stu-id="1a795-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1a795-213">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="1a795-213">Properties:</span></span>

|<span data-ttu-id="1a795-214">Nom</span><span class="sxs-lookup"><span data-stu-id="1a795-214">Name</span></span>| <span data-ttu-id="1a795-215">Type</span><span class="sxs-lookup"><span data-stu-id="1a795-215">Type</span></span>| <span data-ttu-id="1a795-216">Description</span><span class="sxs-lookup"><span data-stu-id="1a795-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1a795-217">Chaîne</span><span class="sxs-lookup"><span data-stu-id="1a795-217">String</span></span>|<span data-ttu-id="1a795-218">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="1a795-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1a795-219">String</span><span class="sxs-lookup"><span data-stu-id="1a795-219">String</span></span>|<span data-ttu-id="1a795-220">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="1a795-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1a795-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1a795-221">Requirements</span></span>

|<span data-ttu-id="1a795-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1a795-222">Requirement</span></span>| <span data-ttu-id="1a795-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="1a795-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="1a795-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1a795-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1a795-225">1.1</span><span class="sxs-lookup"><span data-stu-id="1a795-225">1.1</span></span>|
|[<span data-ttu-id="1a795-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1a795-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1a795-227">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1a795-227">Compose or Read</span></span>|
