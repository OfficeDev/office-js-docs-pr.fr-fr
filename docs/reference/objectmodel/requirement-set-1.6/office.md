---
title: Espace de noms Office-ensemble de conditions requises 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 0a6360ff7f4e397b878d9a3f744bdbe58347c558
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163660"
---
# <a name="office"></a><span data-ttu-id="3fd94-102">Office</span><span class="sxs-lookup"><span data-stu-id="3fd94-102">Office</span></span>

<span data-ttu-id="3fd94-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="3fd94-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="3fd94-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3fd94-105">Requirements</span></span>

|<span data-ttu-id="3fd94-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3fd94-106">Requirement</span></span>| <span data-ttu-id="3fd94-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="3fd94-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fd94-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3fd94-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fd94-109">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-109">1.1</span></span>|
|[<span data-ttu-id="3fd94-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3fd94-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fd94-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3fd94-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="3fd94-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="3fd94-112">Properties</span></span>

| <span data-ttu-id="3fd94-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="3fd94-113">Property</span></span> | <span data-ttu-id="3fd94-114">Modes</span><span class="sxs-lookup"><span data-stu-id="3fd94-114">Modes</span></span> | <span data-ttu-id="3fd94-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="3fd94-115">Return type</span></span> | <span data-ttu-id="3fd94-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="3fd94-116">Minimum</span></span><br><span data-ttu-id="3fd94-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="3fd94-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3fd94-118">context</span><span class="sxs-lookup"><span data-stu-id="3fd94-118">context</span></span>](office.context.md) | <span data-ttu-id="3fd94-119">Composition</span><span class="sxs-lookup"><span data-stu-id="3fd94-119">Compose</span></span><br><span data-ttu-id="3fd94-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="3fd94-120">Read</span></span> | [<span data-ttu-id="3fd94-121">Context</span><span class="sxs-lookup"><span data-stu-id="3fd94-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="3fd94-122">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="3fd94-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="3fd94-123">Enumerations</span></span>

| <span data-ttu-id="3fd94-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="3fd94-124">Enumeration</span></span> | <span data-ttu-id="3fd94-125">Modes</span><span class="sxs-lookup"><span data-stu-id="3fd94-125">Modes</span></span> | <span data-ttu-id="3fd94-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="3fd94-126">Return type</span></span> | <span data-ttu-id="3fd94-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="3fd94-127">Minimum</span></span><br><span data-ttu-id="3fd94-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="3fd94-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="3fd94-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="3fd94-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="3fd94-130">Composition</span><span class="sxs-lookup"><span data-stu-id="3fd94-130">Compose</span></span><br><span data-ttu-id="3fd94-131">Lire</span><span class="sxs-lookup"><span data-stu-id="3fd94-131">Read</span></span> | <span data-ttu-id="3fd94-132">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-132">String</span></span> | [<span data-ttu-id="3fd94-133">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fd94-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="3fd94-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="3fd94-135">Composition</span><span class="sxs-lookup"><span data-stu-id="3fd94-135">Compose</span></span><br><span data-ttu-id="3fd94-136">Lire</span><span class="sxs-lookup"><span data-stu-id="3fd94-136">Read</span></span> | <span data-ttu-id="3fd94-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-137">String</span></span> | [<span data-ttu-id="3fd94-138">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="3fd94-139">EventType</span><span class="sxs-lookup"><span data-stu-id="3fd94-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="3fd94-140">Composition</span><span class="sxs-lookup"><span data-stu-id="3fd94-140">Compose</span></span><br><span data-ttu-id="3fd94-141">Lire</span><span class="sxs-lookup"><span data-stu-id="3fd94-141">Read</span></span> | <span data-ttu-id="3fd94-142">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-142">String</span></span> | [<span data-ttu-id="3fd94-143">1,5</span><span class="sxs-lookup"><span data-stu-id="3fd94-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="3fd94-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="3fd94-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="3fd94-145">Composition</span><span class="sxs-lookup"><span data-stu-id="3fd94-145">Compose</span></span><br><span data-ttu-id="3fd94-146">Lire</span><span class="sxs-lookup"><span data-stu-id="3fd94-146">Read</span></span> | <span data-ttu-id="3fd94-147">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-147">String</span></span> | [<span data-ttu-id="3fd94-148">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="3fd94-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="3fd94-149">Namespaces</span></span>

<span data-ttu-id="3fd94-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="3fd94-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="3fd94-151">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="3fd94-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="3fd94-152">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="3fd94-153">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="3fd94-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="3fd94-154">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-154">Type</span></span>

*   <span data-ttu-id="3fd94-155">String</span><span class="sxs-lookup"><span data-stu-id="3fd94-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3fd94-156">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="3fd94-156">Properties:</span></span>

|<span data-ttu-id="3fd94-157">Nom</span><span class="sxs-lookup"><span data-stu-id="3fd94-157">Name</span></span>| <span data-ttu-id="3fd94-158">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-158">Type</span></span>| <span data-ttu-id="3fd94-159">Description</span><span class="sxs-lookup"><span data-stu-id="3fd94-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="3fd94-160">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-160">String</span></span>|<span data-ttu-id="3fd94-161">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="3fd94-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="3fd94-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-162">String</span></span>|<span data-ttu-id="3fd94-163">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="3fd94-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3fd94-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3fd94-164">Requirements</span></span>

|<span data-ttu-id="3fd94-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3fd94-165">Requirement</span></span>| <span data-ttu-id="3fd94-166">Valeur</span><span class="sxs-lookup"><span data-stu-id="3fd94-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fd94-167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3fd94-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fd94-168">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-168">1.1</span></span>|
|[<span data-ttu-id="3fd94-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3fd94-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fd94-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3fd94-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="3fd94-171">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-171">CoercionType: String</span></span>

<span data-ttu-id="3fd94-172">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="3fd94-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3fd94-173">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-173">Type</span></span>

*   <span data-ttu-id="3fd94-174">String</span><span class="sxs-lookup"><span data-stu-id="3fd94-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3fd94-175">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="3fd94-175">Properties:</span></span>

|<span data-ttu-id="3fd94-176">Nom</span><span class="sxs-lookup"><span data-stu-id="3fd94-176">Name</span></span>| <span data-ttu-id="3fd94-177">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-177">Type</span></span>| <span data-ttu-id="3fd94-178">Description</span><span class="sxs-lookup"><span data-stu-id="3fd94-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="3fd94-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-179">String</span></span>|<span data-ttu-id="3fd94-180">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="3fd94-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="3fd94-181">String</span><span class="sxs-lookup"><span data-stu-id="3fd94-181">String</span></span>|<span data-ttu-id="3fd94-182">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="3fd94-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3fd94-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3fd94-183">Requirements</span></span>

|<span data-ttu-id="3fd94-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3fd94-184">Requirement</span></span>| <span data-ttu-id="3fd94-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="3fd94-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fd94-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3fd94-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fd94-187">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-187">1.1</span></span>|
|[<span data-ttu-id="3fd94-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3fd94-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fd94-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3fd94-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="3fd94-190">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-190">EventType: String</span></span>

<span data-ttu-id="3fd94-191">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="3fd94-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="3fd94-192">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-192">Type</span></span>

*   <span data-ttu-id="3fd94-193">String</span><span class="sxs-lookup"><span data-stu-id="3fd94-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3fd94-194">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="3fd94-194">Properties:</span></span>

| <span data-ttu-id="3fd94-195">Nom</span><span class="sxs-lookup"><span data-stu-id="3fd94-195">Name</span></span> | <span data-ttu-id="3fd94-196">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-196">Type</span></span> | <span data-ttu-id="3fd94-197">Description</span><span class="sxs-lookup"><span data-stu-id="3fd94-197">Description</span></span> | <span data-ttu-id="3fd94-198">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="3fd94-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="3fd94-199">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-199">String</span></span> | <span data-ttu-id="3fd94-200">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="3fd94-200">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="3fd94-201">1,5</span><span class="sxs-lookup"><span data-stu-id="3fd94-201">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3fd94-202">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3fd94-202">Requirements</span></span>

|<span data-ttu-id="3fd94-203">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3fd94-203">Requirement</span></span>| <span data-ttu-id="3fd94-204">Valeur</span><span class="sxs-lookup"><span data-stu-id="3fd94-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fd94-205">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3fd94-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fd94-206">1,5</span><span class="sxs-lookup"><span data-stu-id="3fd94-206">1.5</span></span> |
|[<span data-ttu-id="3fd94-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3fd94-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fd94-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3fd94-208">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="3fd94-209">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-209">SourceProperty: String</span></span>

<span data-ttu-id="3fd94-210">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="3fd94-210">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="3fd94-211">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-211">Type</span></span>

*   <span data-ttu-id="3fd94-212">String</span><span class="sxs-lookup"><span data-stu-id="3fd94-212">String</span></span>

##### <a name="properties"></a><span data-ttu-id="3fd94-213">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="3fd94-213">Properties:</span></span>

|<span data-ttu-id="3fd94-214">Nom</span><span class="sxs-lookup"><span data-stu-id="3fd94-214">Name</span></span>| <span data-ttu-id="3fd94-215">Type</span><span class="sxs-lookup"><span data-stu-id="3fd94-215">Type</span></span>| <span data-ttu-id="3fd94-216">Description</span><span class="sxs-lookup"><span data-stu-id="3fd94-216">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="3fd94-217">Chaîne</span><span class="sxs-lookup"><span data-stu-id="3fd94-217">String</span></span>|<span data-ttu-id="3fd94-218">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="3fd94-218">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="3fd94-219">String</span><span class="sxs-lookup"><span data-stu-id="3fd94-219">String</span></span>|<span data-ttu-id="3fd94-220">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="3fd94-220">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3fd94-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="3fd94-221">Requirements</span></span>

|<span data-ttu-id="3fd94-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="3fd94-222">Requirement</span></span>| <span data-ttu-id="3fd94-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="3fd94-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="3fd94-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="3fd94-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="3fd94-225">1.1</span><span class="sxs-lookup"><span data-stu-id="3fd94-225">1.1</span></span>|
|[<span data-ttu-id="3fd94-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="3fd94-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="3fd94-227">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="3fd94-227">Compose or Read</span></span>|
