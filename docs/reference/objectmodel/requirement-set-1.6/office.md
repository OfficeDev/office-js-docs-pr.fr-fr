---
title: Espace de noms Office-ensemble de conditions requises 1,6
description: Modèle objet pour l’espace de noms de niveau supérieur de l’API des compléments Outlook (version 1,6 de l’API de boîte aux lettres).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ae2f863e054016636ebffc3ff3925cee018036a1
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717648"
---
# <a name="office"></a><span data-ttu-id="d7dfe-103">Office</span><span class="sxs-lookup"><span data-stu-id="d7dfe-103">Office</span></span>

<span data-ttu-id="d7dfe-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d7dfe-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7dfe-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7dfe-106">Requirements</span></span>

|<span data-ttu-id="d7dfe-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7dfe-107">Requirement</span></span>| <span data-ttu-id="d7dfe-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7dfe-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7dfe-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7dfe-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7dfe-110">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-110">1.1</span></span>|
|[<span data-ttu-id="d7dfe-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7dfe-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7dfe-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d7dfe-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="d7dfe-113">Properties</span></span>

| <span data-ttu-id="d7dfe-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="d7dfe-114">Property</span></span> | <span data-ttu-id="d7dfe-115">Modes</span><span class="sxs-lookup"><span data-stu-id="d7dfe-115">Modes</span></span> | <span data-ttu-id="d7dfe-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="d7dfe-116">Return type</span></span> | <span data-ttu-id="d7dfe-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="d7dfe-117">Minimum</span></span><br><span data-ttu-id="d7dfe-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7dfe-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d7dfe-119">context</span><span class="sxs-lookup"><span data-stu-id="d7dfe-119">context</span></span>](office.context.md) | <span data-ttu-id="d7dfe-120">Composition</span><span class="sxs-lookup"><span data-stu-id="d7dfe-120">Compose</span></span><br><span data-ttu-id="d7dfe-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-121">Read</span></span> | [<span data-ttu-id="d7dfe-122">Context</span><span class="sxs-lookup"><span data-stu-id="d7dfe-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="d7dfe-123">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="d7dfe-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="d7dfe-124">Enumerations</span></span>

| <span data-ttu-id="d7dfe-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="d7dfe-125">Enumeration</span></span> | <span data-ttu-id="d7dfe-126">Modes</span><span class="sxs-lookup"><span data-stu-id="d7dfe-126">Modes</span></span> | <span data-ttu-id="d7dfe-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="d7dfe-127">Return type</span></span> | <span data-ttu-id="d7dfe-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="d7dfe-128">Minimum</span></span><br><span data-ttu-id="d7dfe-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7dfe-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d7dfe-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d7dfe-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d7dfe-131">Composition</span><span class="sxs-lookup"><span data-stu-id="d7dfe-131">Compose</span></span><br><span data-ttu-id="d7dfe-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-132">Read</span></span> | <span data-ttu-id="d7dfe-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-133">String</span></span> | [<span data-ttu-id="d7dfe-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7dfe-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d7dfe-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d7dfe-136">Composition</span><span class="sxs-lookup"><span data-stu-id="d7dfe-136">Compose</span></span><br><span data-ttu-id="d7dfe-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-137">Read</span></span> | <span data-ttu-id="d7dfe-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-138">String</span></span> | [<span data-ttu-id="d7dfe-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7dfe-140">EventType</span><span class="sxs-lookup"><span data-stu-id="d7dfe-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d7dfe-141">Composition</span><span class="sxs-lookup"><span data-stu-id="d7dfe-141">Compose</span></span><br><span data-ttu-id="d7dfe-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-142">Read</span></span> | <span data-ttu-id="d7dfe-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-143">String</span></span> | [<span data-ttu-id="d7dfe-144">1,5</span><span class="sxs-lookup"><span data-stu-id="d7dfe-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="d7dfe-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d7dfe-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d7dfe-146">Composition</span><span class="sxs-lookup"><span data-stu-id="d7dfe-146">Compose</span></span><br><span data-ttu-id="d7dfe-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-147">Read</span></span> | <span data-ttu-id="d7dfe-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-148">String</span></span> | [<span data-ttu-id="d7dfe-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="d7dfe-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="d7dfe-150">Namespaces</span></span>

<span data-ttu-id="d7dfe-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d7dfe-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="d7dfe-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d7dfe-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="d7dfe-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d7dfe-155">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-155">Type</span></span>

*   <span data-ttu-id="d7dfe-156">String</span><span class="sxs-lookup"><span data-stu-id="d7dfe-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7dfe-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d7dfe-157">Properties:</span></span>

|<span data-ttu-id="d7dfe-158">Nom</span><span class="sxs-lookup"><span data-stu-id="d7dfe-158">Name</span></span>| <span data-ttu-id="d7dfe-159">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-159">Type</span></span>| <span data-ttu-id="d7dfe-160">Description</span><span class="sxs-lookup"><span data-stu-id="d7dfe-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d7dfe-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-161">String</span></span>|<span data-ttu-id="d7dfe-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d7dfe-163">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-163">String</span></span>|<span data-ttu-id="d7dfe-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7dfe-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7dfe-165">Requirements</span></span>

|<span data-ttu-id="d7dfe-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7dfe-166">Requirement</span></span>| <span data-ttu-id="d7dfe-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7dfe-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7dfe-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7dfe-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7dfe-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-169">1.1</span></span>|
|[<span data-ttu-id="d7dfe-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7dfe-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7dfe-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d7dfe-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-172">CoercionType: String</span></span>

<span data-ttu-id="d7dfe-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7dfe-174">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-174">Type</span></span>

*   <span data-ttu-id="d7dfe-175">String</span><span class="sxs-lookup"><span data-stu-id="d7dfe-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7dfe-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d7dfe-176">Properties:</span></span>

|<span data-ttu-id="d7dfe-177">Nom</span><span class="sxs-lookup"><span data-stu-id="d7dfe-177">Name</span></span>| <span data-ttu-id="d7dfe-178">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-178">Type</span></span>| <span data-ttu-id="d7dfe-179">Description</span><span class="sxs-lookup"><span data-stu-id="d7dfe-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d7dfe-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-180">String</span></span>|<span data-ttu-id="d7dfe-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d7dfe-182">String</span><span class="sxs-lookup"><span data-stu-id="d7dfe-182">String</span></span>|<span data-ttu-id="d7dfe-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7dfe-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7dfe-184">Requirements</span></span>

|<span data-ttu-id="d7dfe-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7dfe-185">Requirement</span></span>| <span data-ttu-id="d7dfe-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7dfe-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7dfe-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7dfe-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7dfe-188">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-188">1.1</span></span>|
|[<span data-ttu-id="d7dfe-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7dfe-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7dfe-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="d7dfe-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-191">EventType: String</span></span>

<span data-ttu-id="d7dfe-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d7dfe-193">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-193">Type</span></span>

*   <span data-ttu-id="d7dfe-194">String</span><span class="sxs-lookup"><span data-stu-id="d7dfe-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7dfe-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d7dfe-195">Properties:</span></span>

| <span data-ttu-id="d7dfe-196">Nom</span><span class="sxs-lookup"><span data-stu-id="d7dfe-196">Name</span></span> | <span data-ttu-id="d7dfe-197">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-197">Type</span></span> | <span data-ttu-id="d7dfe-198">Description</span><span class="sxs-lookup"><span data-stu-id="d7dfe-198">Description</span></span> | <span data-ttu-id="d7dfe-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="d7dfe-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="d7dfe-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-200">String</span></span> | <span data-ttu-id="d7dfe-201">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="d7dfe-202">1,5</span><span class="sxs-lookup"><span data-stu-id="d7dfe-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d7dfe-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7dfe-203">Requirements</span></span>

|<span data-ttu-id="d7dfe-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7dfe-204">Requirement</span></span>| <span data-ttu-id="d7dfe-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7dfe-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7dfe-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7dfe-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7dfe-207">1,5</span><span class="sxs-lookup"><span data-stu-id="d7dfe-207">1.5</span></span> |
|[<span data-ttu-id="d7dfe-208">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7dfe-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7dfe-209">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d7dfe-210">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-210">SourceProperty: String</span></span>

<span data-ttu-id="d7dfe-211">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7dfe-212">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-212">Type</span></span>

*   <span data-ttu-id="d7dfe-213">String</span><span class="sxs-lookup"><span data-stu-id="d7dfe-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7dfe-214">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="d7dfe-214">Properties:</span></span>

|<span data-ttu-id="d7dfe-215">Nom</span><span class="sxs-lookup"><span data-stu-id="d7dfe-215">Name</span></span>| <span data-ttu-id="d7dfe-216">Type</span><span class="sxs-lookup"><span data-stu-id="d7dfe-216">Type</span></span>| <span data-ttu-id="d7dfe-217">Description</span><span class="sxs-lookup"><span data-stu-id="d7dfe-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d7dfe-218">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7dfe-218">String</span></span>|<span data-ttu-id="d7dfe-219">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d7dfe-220">String</span><span class="sxs-lookup"><span data-stu-id="d7dfe-220">String</span></span>|<span data-ttu-id="d7dfe-221">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="d7dfe-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7dfe-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7dfe-222">Requirements</span></span>

|<span data-ttu-id="d7dfe-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7dfe-223">Requirement</span></span>| <span data-ttu-id="d7dfe-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7dfe-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7dfe-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7dfe-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7dfe-226">1.1</span><span class="sxs-lookup"><span data-stu-id="d7dfe-226">1.1</span></span>|
|[<span data-ttu-id="d7dfe-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7dfe-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7dfe-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7dfe-228">Compose or Read</span></span>|
