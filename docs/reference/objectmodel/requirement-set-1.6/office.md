---
title: Espace de noms Office-ensemble de conditions requises 1,6
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: b0d1643727055c6b7ddb4d03c0488b82b24f3fad
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611455"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="dd0d2-103">Office (boîte aux lettres requise définie sur 1,6)</span><span class="sxs-lookup"><span data-stu-id="dd0d2-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="dd0d2-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="dd0d2-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dd0d2-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd0d2-106">Requirements</span></span>

|<span data-ttu-id="dd0d2-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd0d2-107">Requirement</span></span>| <span data-ttu-id="dd0d2-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd0d2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0d2-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd0d2-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0d2-110">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-110">1.1</span></span>|
|[<span data-ttu-id="dd0d2-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd0d2-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0d2-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd0d2-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dd0d2-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="dd0d2-113">Properties</span></span>

| <span data-ttu-id="dd0d2-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="dd0d2-114">Property</span></span> | <span data-ttu-id="dd0d2-115">Modes</span><span class="sxs-lookup"><span data-stu-id="dd0d2-115">Modes</span></span> | <span data-ttu-id="dd0d2-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dd0d2-116">Return type</span></span> | <span data-ttu-id="dd0d2-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="dd0d2-117">Minimum</span></span><br><span data-ttu-id="dd0d2-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd0d2-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dd0d2-119">context</span><span class="sxs-lookup"><span data-stu-id="dd0d2-119">context</span></span>](office.context.md) | <span data-ttu-id="dd0d2-120">Composition</span><span class="sxs-lookup"><span data-stu-id="dd0d2-120">Compose</span></span><br><span data-ttu-id="dd0d2-121">Read</span><span class="sxs-lookup"><span data-stu-id="dd0d2-121">Read</span></span> | [<span data-ttu-id="dd0d2-122">Context</span><span class="sxs-lookup"><span data-stu-id="dd0d2-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6) | [<span data-ttu-id="dd0d2-123">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="dd0d2-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="dd0d2-124">Enumerations</span></span>

| <span data-ttu-id="dd0d2-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="dd0d2-125">Enumeration</span></span> | <span data-ttu-id="dd0d2-126">Modes</span><span class="sxs-lookup"><span data-stu-id="dd0d2-126">Modes</span></span> | <span data-ttu-id="dd0d2-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dd0d2-127">Return type</span></span> | <span data-ttu-id="dd0d2-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="dd0d2-128">Minimum</span></span><br><span data-ttu-id="dd0d2-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd0d2-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dd0d2-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="dd0d2-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="dd0d2-131">Composition</span><span class="sxs-lookup"><span data-stu-id="dd0d2-131">Compose</span></span><br><span data-ttu-id="dd0d2-132">Read</span><span class="sxs-lookup"><span data-stu-id="dd0d2-132">Read</span></span> | <span data-ttu-id="dd0d2-133">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-133">String</span></span> | [<span data-ttu-id="dd0d2-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd0d2-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="dd0d2-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="dd0d2-136">Composition</span><span class="sxs-lookup"><span data-stu-id="dd0d2-136">Compose</span></span><br><span data-ttu-id="dd0d2-137">Read</span><span class="sxs-lookup"><span data-stu-id="dd0d2-137">Read</span></span> | <span data-ttu-id="dd0d2-138">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-138">String</span></span> | [<span data-ttu-id="dd0d2-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dd0d2-140">EventType</span><span class="sxs-lookup"><span data-stu-id="dd0d2-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="dd0d2-141">Composition</span><span class="sxs-lookup"><span data-stu-id="dd0d2-141">Compose</span></span><br><span data-ttu-id="dd0d2-142">Read</span><span class="sxs-lookup"><span data-stu-id="dd0d2-142">Read</span></span> | <span data-ttu-id="dd0d2-143">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-143">String</span></span> | [<span data-ttu-id="dd0d2-144">1,5</span><span class="sxs-lookup"><span data-stu-id="dd0d2-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="dd0d2-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="dd0d2-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="dd0d2-146">Composition</span><span class="sxs-lookup"><span data-stu-id="dd0d2-146">Compose</span></span><br><span data-ttu-id="dd0d2-147">Read</span><span class="sxs-lookup"><span data-stu-id="dd0d2-147">Read</span></span> | <span data-ttu-id="dd0d2-148">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-148">String</span></span> | [<span data-ttu-id="dd0d2-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="dd0d2-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="dd0d2-150">Namespaces</span></span>

<span data-ttu-id="dd0d2-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="dd0d2-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="dd0d2-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="dd0d2-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="dd0d2-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="dd0d2-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="dd0d2-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0d2-155">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-155">Type</span></span>

*   <span data-ttu-id="dd0d2-156">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dd0d2-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dd0d2-157">Properties:</span></span>

|<span data-ttu-id="dd0d2-158">Nom</span><span class="sxs-lookup"><span data-stu-id="dd0d2-158">Name</span></span>| <span data-ttu-id="dd0d2-159">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-159">Type</span></span>| <span data-ttu-id="dd0d2-160">Description</span><span class="sxs-lookup"><span data-stu-id="dd0d2-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="dd0d2-161">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-161">String</span></span>|<span data-ttu-id="dd0d2-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="dd0d2-163">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-163">String</span></span>|<span data-ttu-id="dd0d2-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0d2-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd0d2-165">Requirements</span></span>

|<span data-ttu-id="dd0d2-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd0d2-166">Requirement</span></span>| <span data-ttu-id="dd0d2-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd0d2-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0d2-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd0d2-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0d2-169">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-169">1.1</span></span>|
|[<span data-ttu-id="dd0d2-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd0d2-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0d2-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd0d2-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="dd0d2-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="dd0d2-172">CoercionType: String</span></span>

<span data-ttu-id="dd0d2-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0d2-174">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-174">Type</span></span>

*   <span data-ttu-id="dd0d2-175">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dd0d2-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dd0d2-176">Properties:</span></span>

|<span data-ttu-id="dd0d2-177">Nom</span><span class="sxs-lookup"><span data-stu-id="dd0d2-177">Name</span></span>| <span data-ttu-id="dd0d2-178">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-178">Type</span></span>| <span data-ttu-id="dd0d2-179">Description</span><span class="sxs-lookup"><span data-stu-id="dd0d2-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="dd0d2-180">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-180">String</span></span>|<span data-ttu-id="dd0d2-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="dd0d2-182">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-182">String</span></span>|<span data-ttu-id="dd0d2-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0d2-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd0d2-184">Requirements</span></span>

|<span data-ttu-id="dd0d2-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd0d2-185">Requirement</span></span>| <span data-ttu-id="dd0d2-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd0d2-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0d2-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd0d2-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0d2-188">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-188">1.1</span></span>|
|[<span data-ttu-id="dd0d2-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd0d2-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0d2-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd0d2-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="dd0d2-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="dd0d2-191">EventType: String</span></span>

<span data-ttu-id="dd0d2-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0d2-193">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-193">Type</span></span>

*   <span data-ttu-id="dd0d2-194">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dd0d2-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dd0d2-195">Properties:</span></span>

| <span data-ttu-id="dd0d2-196">Nom</span><span class="sxs-lookup"><span data-stu-id="dd0d2-196">Name</span></span> | <span data-ttu-id="dd0d2-197">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-197">Type</span></span> | <span data-ttu-id="dd0d2-198">Description</span><span class="sxs-lookup"><span data-stu-id="dd0d2-198">Description</span></span> | <span data-ttu-id="dd0d2-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="dd0d2-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="dd0d2-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dd0d2-200">String</span></span> | <span data-ttu-id="dd0d2-201">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="dd0d2-202">1,5</span><span class="sxs-lookup"><span data-stu-id="dd0d2-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="dd0d2-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd0d2-203">Requirements</span></span>

|<span data-ttu-id="dd0d2-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd0d2-204">Requirement</span></span>| <span data-ttu-id="dd0d2-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd0d2-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0d2-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd0d2-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0d2-207">1,5</span><span class="sxs-lookup"><span data-stu-id="dd0d2-207">1.5</span></span> |
|[<span data-ttu-id="dd0d2-208">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd0d2-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0d2-209">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd0d2-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="dd0d2-210">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="dd0d2-210">SourceProperty: String</span></span>

<span data-ttu-id="dd0d2-211">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="dd0d2-212">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-212">Type</span></span>

*   <span data-ttu-id="dd0d2-213">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="dd0d2-214">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="dd0d2-214">Properties:</span></span>

|<span data-ttu-id="dd0d2-215">Nom</span><span class="sxs-lookup"><span data-stu-id="dd0d2-215">Name</span></span>| <span data-ttu-id="dd0d2-216">Type</span><span class="sxs-lookup"><span data-stu-id="dd0d2-216">Type</span></span>| <span data-ttu-id="dd0d2-217">Description</span><span class="sxs-lookup"><span data-stu-id="dd0d2-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="dd0d2-218">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-218">String</span></span>|<span data-ttu-id="dd0d2-219">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="dd0d2-220">String</span><span class="sxs-lookup"><span data-stu-id="dd0d2-220">String</span></span>|<span data-ttu-id="dd0d2-221">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="dd0d2-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="dd0d2-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dd0d2-222">Requirements</span></span>

|<span data-ttu-id="dd0d2-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dd0d2-223">Requirement</span></span>| <span data-ttu-id="dd0d2-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="dd0d2-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="dd0d2-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dd0d2-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dd0d2-226">1.1</span><span class="sxs-lookup"><span data-stu-id="dd0d2-226">1.1</span></span>|
|[<span data-ttu-id="dd0d2-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dd0d2-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dd0d2-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dd0d2-228">Compose or Read</span></span>|
