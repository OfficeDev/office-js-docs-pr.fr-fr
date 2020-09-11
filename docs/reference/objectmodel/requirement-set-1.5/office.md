---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 141fd124ba5778a5ae576c7b4cd2c749a9c4bd6f
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430596"
---
# <a name="office-mailbox-requirement-set-15"></a><span data-ttu-id="cd9a4-103">Office (boîte aux lettres requise définie sur 1,5)</span><span class="sxs-lookup"><span data-stu-id="cd9a4-103">Office (Mailbox requirement set 1.5)</span></span>

<span data-ttu-id="cd9a4-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="cd9a4-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cd9a4-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cd9a4-106">Requirements</span></span>

|<span data-ttu-id="cd9a4-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cd9a4-107">Requirement</span></span>| <span data-ttu-id="cd9a4-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="cd9a4-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd9a4-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cd9a4-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd9a4-110">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-110">1.1</span></span>|
|[<span data-ttu-id="cd9a4-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cd9a4-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd9a4-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="cd9a4-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="cd9a4-113">Properties</span></span>

| <span data-ttu-id="cd9a4-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="cd9a4-114">Property</span></span> | <span data-ttu-id="cd9a4-115">Modes</span><span class="sxs-lookup"><span data-stu-id="cd9a4-115">Modes</span></span> | <span data-ttu-id="cd9a4-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="cd9a4-116">Return type</span></span> | <span data-ttu-id="cd9a4-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="cd9a4-117">Minimum</span></span><br><span data-ttu-id="cd9a4-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="cd9a4-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cd9a4-119">context</span><span class="sxs-lookup"><span data-stu-id="cd9a4-119">context</span></span>](office.context.md) | <span data-ttu-id="cd9a4-120">Composition</span><span class="sxs-lookup"><span data-stu-id="cd9a4-120">Compose</span></span><br><span data-ttu-id="cd9a4-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-121">Read</span></span> | [<span data-ttu-id="cd9a4-122">Context</span><span class="sxs-lookup"><span data-stu-id="cd9a4-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="cd9a4-123">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="cd9a4-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="cd9a4-124">Enumerations</span></span>

| <span data-ttu-id="cd9a4-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="cd9a4-125">Enumeration</span></span> | <span data-ttu-id="cd9a4-126">Modes</span><span class="sxs-lookup"><span data-stu-id="cd9a4-126">Modes</span></span> | <span data-ttu-id="cd9a4-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="cd9a4-127">Return type</span></span> | <span data-ttu-id="cd9a4-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="cd9a4-128">Minimum</span></span><br><span data-ttu-id="cd9a4-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="cd9a4-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cd9a4-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="cd9a4-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="cd9a4-131">Composition</span><span class="sxs-lookup"><span data-stu-id="cd9a4-131">Compose</span></span><br><span data-ttu-id="cd9a4-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-132">Read</span></span> | <span data-ttu-id="cd9a4-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-133">String</span></span> | [<span data-ttu-id="cd9a4-134">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cd9a4-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="cd9a4-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="cd9a4-136">Composition</span><span class="sxs-lookup"><span data-stu-id="cd9a4-136">Compose</span></span><br><span data-ttu-id="cd9a4-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-137">Read</span></span> | <span data-ttu-id="cd9a4-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-138">String</span></span> | [<span data-ttu-id="cd9a4-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cd9a4-140">EventType</span><span class="sxs-lookup"><span data-stu-id="cd9a4-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="cd9a4-141">Composition</span><span class="sxs-lookup"><span data-stu-id="cd9a4-141">Compose</span></span><br><span data-ttu-id="cd9a4-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-142">Read</span></span> | <span data-ttu-id="cd9a4-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-143">String</span></span> | [<span data-ttu-id="cd9a4-144">1,5</span><span class="sxs-lookup"><span data-stu-id="cd9a4-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="cd9a4-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="cd9a4-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="cd9a4-146">Composition</span><span class="sxs-lookup"><span data-stu-id="cd9a4-146">Compose</span></span><br><span data-ttu-id="cd9a4-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-147">Read</span></span> | <span data-ttu-id="cd9a4-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-148">String</span></span> | [<span data-ttu-id="cd9a4-149">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="cd9a4-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="cd9a4-150">Namespaces</span></span>

<span data-ttu-id="cd9a4-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="cd9a4-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.5&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="cd9a4-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="cd9a4-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="cd9a4-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="cd9a4-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="cd9a4-155">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-155">Type</span></span>

*   <span data-ttu-id="cd9a4-156">String</span><span class="sxs-lookup"><span data-stu-id="cd9a4-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd9a4-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="cd9a4-157">Properties:</span></span>

|<span data-ttu-id="cd9a4-158">Nom</span><span class="sxs-lookup"><span data-stu-id="cd9a4-158">Name</span></span>| <span data-ttu-id="cd9a4-159">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-159">Type</span></span>| <span data-ttu-id="cd9a4-160">Description</span><span class="sxs-lookup"><span data-stu-id="cd9a4-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="cd9a4-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-161">String</span></span>|<span data-ttu-id="cd9a4-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="cd9a4-163">String</span><span class="sxs-lookup"><span data-stu-id="cd9a4-163">String</span></span>|<span data-ttu-id="cd9a4-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cd9a4-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cd9a4-165">Requirements</span></span>

|<span data-ttu-id="cd9a4-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cd9a4-166">Requirement</span></span>| <span data-ttu-id="cd9a4-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="cd9a4-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd9a4-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cd9a4-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd9a4-169">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-169">1.1</span></span>|
|[<span data-ttu-id="cd9a4-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cd9a4-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd9a4-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="cd9a4-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-172">CoercionType: String</span></span>

<span data-ttu-id="cd9a4-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cd9a4-174">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-174">Type</span></span>

*   <span data-ttu-id="cd9a4-175">String</span><span class="sxs-lookup"><span data-stu-id="cd9a4-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd9a4-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="cd9a4-176">Properties:</span></span>

|<span data-ttu-id="cd9a4-177">Nom</span><span class="sxs-lookup"><span data-stu-id="cd9a4-177">Name</span></span>| <span data-ttu-id="cd9a4-178">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-178">Type</span></span>| <span data-ttu-id="cd9a4-179">Description</span><span class="sxs-lookup"><span data-stu-id="cd9a4-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="cd9a4-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-180">String</span></span>|<span data-ttu-id="cd9a4-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="cd9a4-182">String</span><span class="sxs-lookup"><span data-stu-id="cd9a4-182">String</span></span>|<span data-ttu-id="cd9a4-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cd9a4-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cd9a4-184">Requirements</span></span>

|<span data-ttu-id="cd9a4-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cd9a4-185">Requirement</span></span>| <span data-ttu-id="cd9a4-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="cd9a4-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd9a4-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cd9a4-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd9a4-188">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-188">1.1</span></span>|
|[<span data-ttu-id="cd9a4-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cd9a4-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd9a4-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="cd9a4-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-191">EventType: String</span></span>

<span data-ttu-id="cd9a4-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="cd9a4-193">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-193">Type</span></span>

*   <span data-ttu-id="cd9a4-194">String</span><span class="sxs-lookup"><span data-stu-id="cd9a4-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd9a4-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="cd9a4-195">Properties:</span></span>

| <span data-ttu-id="cd9a4-196">Nom</span><span class="sxs-lookup"><span data-stu-id="cd9a4-196">Name</span></span> | <span data-ttu-id="cd9a4-197">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-197">Type</span></span> | <span data-ttu-id="cd9a4-198">Description</span><span class="sxs-lookup"><span data-stu-id="cd9a4-198">Description</span></span> | <span data-ttu-id="cd9a4-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="cd9a4-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="cd9a4-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-200">String</span></span> | <span data-ttu-id="cd9a4-201">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="cd9a4-202">1,5</span><span class="sxs-lookup"><span data-stu-id="cd9a4-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="cd9a4-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cd9a4-203">Requirements</span></span>

|<span data-ttu-id="cd9a4-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cd9a4-204">Requirement</span></span>| <span data-ttu-id="cd9a4-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="cd9a4-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd9a4-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cd9a4-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd9a4-207">1,5</span><span class="sxs-lookup"><span data-stu-id="cd9a4-207">1.5</span></span> |
|[<span data-ttu-id="cd9a4-208">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cd9a4-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd9a4-209">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="cd9a4-210">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-210">SourceProperty: String</span></span>

<span data-ttu-id="cd9a4-211">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="cd9a4-212">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-212">Type</span></span>

*   <span data-ttu-id="cd9a4-213">String</span><span class="sxs-lookup"><span data-stu-id="cd9a4-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="cd9a4-214">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="cd9a4-214">Properties:</span></span>

|<span data-ttu-id="cd9a4-215">Nom</span><span class="sxs-lookup"><span data-stu-id="cd9a4-215">Name</span></span>| <span data-ttu-id="cd9a4-216">Type</span><span class="sxs-lookup"><span data-stu-id="cd9a4-216">Type</span></span>| <span data-ttu-id="cd9a4-217">Description</span><span class="sxs-lookup"><span data-stu-id="cd9a4-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="cd9a4-218">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cd9a4-218">String</span></span>|<span data-ttu-id="cd9a4-219">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="cd9a4-220">String</span><span class="sxs-lookup"><span data-stu-id="cd9a4-220">String</span></span>|<span data-ttu-id="cd9a4-221">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="cd9a4-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="cd9a4-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cd9a4-222">Requirements</span></span>

|<span data-ttu-id="cd9a4-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cd9a4-223">Requirement</span></span>| <span data-ttu-id="cd9a4-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="cd9a4-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="cd9a4-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cd9a4-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cd9a4-226">1.1</span><span class="sxs-lookup"><span data-stu-id="cd9a4-226">1.1</span></span>|
|[<span data-ttu-id="cd9a4-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cd9a4-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cd9a4-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cd9a4-228">Compose or Read</span></span>|
