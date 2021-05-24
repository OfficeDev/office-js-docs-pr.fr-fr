---
title: Office de noms - ensemble de conditions requises 1.6
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.6.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 40cdb7de0678007b93b9251e7f1e2921ed857338
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590833"
---
# <a name="office-mailbox-requirement-set-16"></a><span data-ttu-id="5d307-103">Office (ensemble de conditions requises de boîte aux lettres 1.6)</span><span class="sxs-lookup"><span data-stu-id="5d307-103">Office (Mailbox requirement set 1.6)</span></span>

<span data-ttu-id="5d307-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="5d307-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5d307-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d307-106">Requirements</span></span>

|<span data-ttu-id="5d307-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d307-107">Requirement</span></span>| <span data-ttu-id="5d307-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d307-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d307-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d307-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5d307-110">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-110">1.1</span></span>|
|[<span data-ttu-id="5d307-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d307-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5d307-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d307-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="5d307-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="5d307-113">Properties</span></span>

| <span data-ttu-id="5d307-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="5d307-114">Property</span></span> | <span data-ttu-id="5d307-115">Modes</span><span class="sxs-lookup"><span data-stu-id="5d307-115">Modes</span></span> | <span data-ttu-id="5d307-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="5d307-116">Return type</span></span> | <span data-ttu-id="5d307-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="5d307-117">Minimum</span></span><br><span data-ttu-id="5d307-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d307-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5d307-119">context</span><span class="sxs-lookup"><span data-stu-id="5d307-119">context</span></span>](office.context.md) | <span data-ttu-id="5d307-120">Composition</span><span class="sxs-lookup"><span data-stu-id="5d307-120">Compose</span></span><br><span data-ttu-id="5d307-121">Lire</span><span class="sxs-lookup"><span data-stu-id="5d307-121">Read</span></span> | [<span data-ttu-id="5d307-122">Context</span><span class="sxs-lookup"><span data-stu-id="5d307-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="5d307-123">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="5d307-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="5d307-124">Enumerations</span></span>

| <span data-ttu-id="5d307-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="5d307-125">Enumeration</span></span> | <span data-ttu-id="5d307-126">Modes</span><span class="sxs-lookup"><span data-stu-id="5d307-126">Modes</span></span> | <span data-ttu-id="5d307-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="5d307-127">Return type</span></span> | <span data-ttu-id="5d307-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="5d307-128">Minimum</span></span><br><span data-ttu-id="5d307-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d307-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5d307-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="5d307-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="5d307-131">Composition</span><span class="sxs-lookup"><span data-stu-id="5d307-131">Compose</span></span><br><span data-ttu-id="5d307-132">Lire</span><span class="sxs-lookup"><span data-stu-id="5d307-132">Read</span></span> | <span data-ttu-id="5d307-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-133">String</span></span> | [<span data-ttu-id="5d307-134">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5d307-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="5d307-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="5d307-136">Composition</span><span class="sxs-lookup"><span data-stu-id="5d307-136">Compose</span></span><br><span data-ttu-id="5d307-137">Lire</span><span class="sxs-lookup"><span data-stu-id="5d307-137">Read</span></span> | <span data-ttu-id="5d307-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-138">String</span></span> | [<span data-ttu-id="5d307-139">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5d307-140">EventType</span><span class="sxs-lookup"><span data-stu-id="5d307-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="5d307-141">Composition</span><span class="sxs-lookup"><span data-stu-id="5d307-141">Compose</span></span><br><span data-ttu-id="5d307-142">Lire</span><span class="sxs-lookup"><span data-stu-id="5d307-142">Read</span></span> | <span data-ttu-id="5d307-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-143">String</span></span> | [<span data-ttu-id="5d307-144">1.5</span><span class="sxs-lookup"><span data-stu-id="5d307-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="5d307-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="5d307-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="5d307-146">Composition</span><span class="sxs-lookup"><span data-stu-id="5d307-146">Compose</span></span><br><span data-ttu-id="5d307-147">Lire</span><span class="sxs-lookup"><span data-stu-id="5d307-147">Read</span></span> | <span data-ttu-id="5d307-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-148">String</span></span> | [<span data-ttu-id="5d307-149">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="5d307-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="5d307-150">Namespaces</span></span>

<span data-ttu-id="5d307-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="5d307-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.6&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="5d307-152">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="5d307-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="5d307-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="5d307-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="5d307-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="5d307-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="5d307-155">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-155">Type</span></span>

*   <span data-ttu-id="5d307-156">String</span><span class="sxs-lookup"><span data-stu-id="5d307-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5d307-157">Propriétés</span><span class="sxs-lookup"><span data-stu-id="5d307-157">Properties</span></span>

|<span data-ttu-id="5d307-158">Nom</span><span class="sxs-lookup"><span data-stu-id="5d307-158">Name</span></span>| <span data-ttu-id="5d307-159">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-159">Type</span></span>| <span data-ttu-id="5d307-160">Description</span><span class="sxs-lookup"><span data-stu-id="5d307-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="5d307-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-161">String</span></span>|<span data-ttu-id="5d307-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="5d307-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="5d307-163">String</span><span class="sxs-lookup"><span data-stu-id="5d307-163">String</span></span>|<span data-ttu-id="5d307-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="5d307-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d307-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d307-165">Requirements</span></span>

|<span data-ttu-id="5d307-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d307-166">Requirement</span></span>| <span data-ttu-id="5d307-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d307-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d307-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d307-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5d307-169">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-169">1.1</span></span>|
|[<span data-ttu-id="5d307-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d307-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5d307-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d307-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="5d307-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="5d307-172">CoercionType: String</span></span>

<span data-ttu-id="5d307-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="5d307-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5d307-174">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-174">Type</span></span>

*   <span data-ttu-id="5d307-175">String</span><span class="sxs-lookup"><span data-stu-id="5d307-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5d307-176">Propriétés</span><span class="sxs-lookup"><span data-stu-id="5d307-176">Properties</span></span>

|<span data-ttu-id="5d307-177">Nom</span><span class="sxs-lookup"><span data-stu-id="5d307-177">Name</span></span>| <span data-ttu-id="5d307-178">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-178">Type</span></span>| <span data-ttu-id="5d307-179">Description</span><span class="sxs-lookup"><span data-stu-id="5d307-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="5d307-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-180">String</span></span>|<span data-ttu-id="5d307-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="5d307-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="5d307-182">String</span><span class="sxs-lookup"><span data-stu-id="5d307-182">String</span></span>|<span data-ttu-id="5d307-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="5d307-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d307-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d307-184">Requirements</span></span>

|<span data-ttu-id="5d307-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d307-185">Requirement</span></span>| <span data-ttu-id="5d307-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d307-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d307-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d307-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5d307-188">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-188">1.1</span></span>|
|[<span data-ttu-id="5d307-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d307-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5d307-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d307-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="5d307-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="5d307-191">EventType: String</span></span>

<span data-ttu-id="5d307-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="5d307-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="5d307-193">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-193">Type</span></span>

*   <span data-ttu-id="5d307-194">String</span><span class="sxs-lookup"><span data-stu-id="5d307-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5d307-195">Propriétés</span><span class="sxs-lookup"><span data-stu-id="5d307-195">Properties</span></span>

| <span data-ttu-id="5d307-196">Nom</span><span class="sxs-lookup"><span data-stu-id="5d307-196">Name</span></span> | <span data-ttu-id="5d307-197">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-197">Type</span></span> | <span data-ttu-id="5d307-198">Description</span><span class="sxs-lookup"><span data-stu-id="5d307-198">Description</span></span> | <span data-ttu-id="5d307-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="5d307-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`ItemChanged`| <span data-ttu-id="5d307-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-200">String</span></span> | <span data-ttu-id="5d307-201">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="5d307-201">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="5d307-202">1,5</span><span class="sxs-lookup"><span data-stu-id="5d307-202">1.5</span></span> |

##### <a name="requirements"></a><span data-ttu-id="5d307-203">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d307-203">Requirements</span></span>

|<span data-ttu-id="5d307-204">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d307-204">Requirement</span></span>| <span data-ttu-id="5d307-205">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d307-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d307-206">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d307-206">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5d307-207">1,5</span><span class="sxs-lookup"><span data-stu-id="5d307-207">1.5</span></span> |
|[<span data-ttu-id="5d307-208">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d307-208">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5d307-209">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d307-209">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="5d307-210">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="5d307-210">SourceProperty: String</span></span>

<span data-ttu-id="5d307-211">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="5d307-211">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="5d307-212">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-212">Type</span></span>

*   <span data-ttu-id="5d307-213">String</span><span class="sxs-lookup"><span data-stu-id="5d307-213">String</span></span>

##### <a name="properties"></a><span data-ttu-id="5d307-214">Propriétés</span><span class="sxs-lookup"><span data-stu-id="5d307-214">Properties</span></span>

|<span data-ttu-id="5d307-215">Nom</span><span class="sxs-lookup"><span data-stu-id="5d307-215">Name</span></span>| <span data-ttu-id="5d307-216">Type</span><span class="sxs-lookup"><span data-stu-id="5d307-216">Type</span></span>| <span data-ttu-id="5d307-217">Description</span><span class="sxs-lookup"><span data-stu-id="5d307-217">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="5d307-218">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5d307-218">String</span></span>|<span data-ttu-id="5d307-219">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="5d307-219">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="5d307-220">String</span><span class="sxs-lookup"><span data-stu-id="5d307-220">String</span></span>|<span data-ttu-id="5d307-221">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="5d307-221">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5d307-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5d307-222">Requirements</span></span>

|<span data-ttu-id="5d307-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5d307-223">Requirement</span></span>| <span data-ttu-id="5d307-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="5d307-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="5d307-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5d307-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5d307-226">1.1</span><span class="sxs-lookup"><span data-stu-id="5d307-226">1.1</span></span>|
|[<span data-ttu-id="5d307-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5d307-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5d307-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5d307-228">Compose or Read</span></span>|
