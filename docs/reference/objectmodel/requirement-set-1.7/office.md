---
title: Espace de noms Office-ensemble de conditions requises 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 23f3fb705c03eabd8ee7fce53f4c89a48128672f
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165347"
---
# <a name="office"></a><span data-ttu-id="2f260-102">Office</span><span class="sxs-lookup"><span data-stu-id="2f260-102">Office</span></span>

<span data-ttu-id="2f260-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="2f260-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2f260-105">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2f260-105">Requirements</span></span>

|<span data-ttu-id="2f260-106">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f260-106">Requirement</span></span>| <span data-ttu-id="2f260-107">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f260-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f260-108">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f260-108">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f260-109">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-109">1.1</span></span>|
|[<span data-ttu-id="2f260-110">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f260-110">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f260-111">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f260-111">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2f260-112">Propriétés</span><span class="sxs-lookup"><span data-stu-id="2f260-112">Properties</span></span>

| <span data-ttu-id="2f260-113">Propriété</span><span class="sxs-lookup"><span data-stu-id="2f260-113">Property</span></span> | <span data-ttu-id="2f260-114">Modes</span><span class="sxs-lookup"><span data-stu-id="2f260-114">Modes</span></span> | <span data-ttu-id="2f260-115">Type de retour</span><span class="sxs-lookup"><span data-stu-id="2f260-115">Return type</span></span> | <span data-ttu-id="2f260-116">Minimale</span><span class="sxs-lookup"><span data-stu-id="2f260-116">Minimum</span></span><br><span data-ttu-id="2f260-117">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f260-117">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2f260-118">context</span><span class="sxs-lookup"><span data-stu-id="2f260-118">context</span></span>](office.context.md) | <span data-ttu-id="2f260-119">Composition</span><span class="sxs-lookup"><span data-stu-id="2f260-119">Compose</span></span><br><span data-ttu-id="2f260-120">Lecture</span><span class="sxs-lookup"><span data-stu-id="2f260-120">Read</span></span> | [<span data-ttu-id="2f260-121">Context</span><span class="sxs-lookup"><span data-stu-id="2f260-121">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="2f260-122">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-122">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="2f260-123">Énumérations</span><span class="sxs-lookup"><span data-stu-id="2f260-123">Enumerations</span></span>

| <span data-ttu-id="2f260-124">Énumération</span><span class="sxs-lookup"><span data-stu-id="2f260-124">Enumeration</span></span> | <span data-ttu-id="2f260-125">Modes</span><span class="sxs-lookup"><span data-stu-id="2f260-125">Modes</span></span> | <span data-ttu-id="2f260-126">Type de retour</span><span class="sxs-lookup"><span data-stu-id="2f260-126">Return type</span></span> | <span data-ttu-id="2f260-127">Minimale</span><span class="sxs-lookup"><span data-stu-id="2f260-127">Minimum</span></span><br><span data-ttu-id="2f260-128">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f260-128">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2f260-129">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="2f260-129">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="2f260-130">Composition</span><span class="sxs-lookup"><span data-stu-id="2f260-130">Compose</span></span><br><span data-ttu-id="2f260-131">Lire</span><span class="sxs-lookup"><span data-stu-id="2f260-131">Read</span></span> | <span data-ttu-id="2f260-132">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-132">String</span></span> | [<span data-ttu-id="2f260-133">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f260-134">CoercionType</span><span class="sxs-lookup"><span data-stu-id="2f260-134">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="2f260-135">Composition</span><span class="sxs-lookup"><span data-stu-id="2f260-135">Compose</span></span><br><span data-ttu-id="2f260-136">Lire</span><span class="sxs-lookup"><span data-stu-id="2f260-136">Read</span></span> | <span data-ttu-id="2f260-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-137">String</span></span> | [<span data-ttu-id="2f260-138">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2f260-139">EventType</span><span class="sxs-lookup"><span data-stu-id="2f260-139">EventType</span></span>](#eventtype-string) | <span data-ttu-id="2f260-140">Composition</span><span class="sxs-lookup"><span data-stu-id="2f260-140">Compose</span></span><br><span data-ttu-id="2f260-141">Lire</span><span class="sxs-lookup"><span data-stu-id="2f260-141">Read</span></span> | <span data-ttu-id="2f260-142">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-142">String</span></span> | [<span data-ttu-id="2f260-143">1,5</span><span class="sxs-lookup"><span data-stu-id="2f260-143">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="2f260-144">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="2f260-144">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="2f260-145">Composition</span><span class="sxs-lookup"><span data-stu-id="2f260-145">Compose</span></span><br><span data-ttu-id="2f260-146">Lire</span><span class="sxs-lookup"><span data-stu-id="2f260-146">Read</span></span> | <span data-ttu-id="2f260-147">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-147">String</span></span> | [<span data-ttu-id="2f260-148">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="2f260-149">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="2f260-149">Namespaces</span></span>

<span data-ttu-id="2f260-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="2f260-150">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="2f260-151">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="2f260-151">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="2f260-152">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-152">AsyncResultStatus: String</span></span>

<span data-ttu-id="2f260-153">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="2f260-153">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="2f260-154">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-154">Type</span></span>

*   <span data-ttu-id="2f260-155">String</span><span class="sxs-lookup"><span data-stu-id="2f260-155">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f260-156">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2f260-156">Properties:</span></span>

|<span data-ttu-id="2f260-157">Nom</span><span class="sxs-lookup"><span data-stu-id="2f260-157">Name</span></span>| <span data-ttu-id="2f260-158">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-158">Type</span></span>| <span data-ttu-id="2f260-159">Description</span><span class="sxs-lookup"><span data-stu-id="2f260-159">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="2f260-160">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-160">String</span></span>|<span data-ttu-id="2f260-161">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="2f260-161">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="2f260-162">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-162">String</span></span>|<span data-ttu-id="2f260-163">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="2f260-163">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f260-164">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2f260-164">Requirements</span></span>

|<span data-ttu-id="2f260-165">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f260-165">Requirement</span></span>| <span data-ttu-id="2f260-166">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f260-166">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f260-167">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f260-167">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f260-168">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-168">1.1</span></span>|
|[<span data-ttu-id="2f260-169">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f260-169">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f260-170">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f260-170">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="2f260-171">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-171">CoercionType: String</span></span>

<span data-ttu-id="2f260-172">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="2f260-172">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2f260-173">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-173">Type</span></span>

*   <span data-ttu-id="2f260-174">String</span><span class="sxs-lookup"><span data-stu-id="2f260-174">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f260-175">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2f260-175">Properties:</span></span>

|<span data-ttu-id="2f260-176">Nom</span><span class="sxs-lookup"><span data-stu-id="2f260-176">Name</span></span>| <span data-ttu-id="2f260-177">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-177">Type</span></span>| <span data-ttu-id="2f260-178">Description</span><span class="sxs-lookup"><span data-stu-id="2f260-178">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="2f260-179">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-179">String</span></span>|<span data-ttu-id="2f260-180">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="2f260-180">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="2f260-181">String</span><span class="sxs-lookup"><span data-stu-id="2f260-181">String</span></span>|<span data-ttu-id="2f260-182">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="2f260-182">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f260-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2f260-183">Requirements</span></span>

|<span data-ttu-id="2f260-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f260-184">Requirement</span></span>| <span data-ttu-id="2f260-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f260-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f260-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f260-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f260-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-187">1.1</span></span>|
|[<span data-ttu-id="2f260-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f260-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f260-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f260-189">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="2f260-190">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-190">EventType: String</span></span>

<span data-ttu-id="2f260-191">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="2f260-191">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="2f260-192">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-192">Type</span></span>

*   <span data-ttu-id="2f260-193">String</span><span class="sxs-lookup"><span data-stu-id="2f260-193">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f260-194">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2f260-194">Properties:</span></span>

| <span data-ttu-id="2f260-195">Nom</span><span class="sxs-lookup"><span data-stu-id="2f260-195">Name</span></span> | <span data-ttu-id="2f260-196">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-196">Type</span></span> | <span data-ttu-id="2f260-197">Description</span><span class="sxs-lookup"><span data-stu-id="2f260-197">Description</span></span> | <span data-ttu-id="2f260-198">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="2f260-198">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="2f260-199">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-199">String</span></span> | <span data-ttu-id="2f260-200">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="2f260-200">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="2f260-201">1.7</span><span class="sxs-lookup"><span data-stu-id="2f260-201">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="2f260-202">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-202">String</span></span> | <span data-ttu-id="2f260-203">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="2f260-203">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="2f260-204">1,5</span><span class="sxs-lookup"><span data-stu-id="2f260-204">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="2f260-205">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-205">String</span></span> | <span data-ttu-id="2f260-206">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="2f260-206">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="2f260-207">1.7</span><span class="sxs-lookup"><span data-stu-id="2f260-207">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="2f260-208">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-208">String</span></span> | <span data-ttu-id="2f260-209">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="2f260-209">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="2f260-210">1.7</span><span class="sxs-lookup"><span data-stu-id="2f260-210">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="2f260-211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2f260-211">Requirements</span></span>

|<span data-ttu-id="2f260-212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f260-212">Requirement</span></span>| <span data-ttu-id="2f260-213">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f260-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f260-214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f260-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f260-215">1,5</span><span class="sxs-lookup"><span data-stu-id="2f260-215">1.5</span></span> |
|[<span data-ttu-id="2f260-216">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f260-216">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f260-217">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f260-217">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="2f260-218">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-218">SourceProperty: String</span></span>

<span data-ttu-id="2f260-219">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="2f260-219">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="2f260-220">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-220">Type</span></span>

*   <span data-ttu-id="2f260-221">String</span><span class="sxs-lookup"><span data-stu-id="2f260-221">String</span></span>

##### <a name="properties"></a><span data-ttu-id="2f260-222">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="2f260-222">Properties:</span></span>

|<span data-ttu-id="2f260-223">Nom</span><span class="sxs-lookup"><span data-stu-id="2f260-223">Name</span></span>| <span data-ttu-id="2f260-224">Type</span><span class="sxs-lookup"><span data-stu-id="2f260-224">Type</span></span>| <span data-ttu-id="2f260-225">Description</span><span class="sxs-lookup"><span data-stu-id="2f260-225">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="2f260-226">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2f260-226">String</span></span>|<span data-ttu-id="2f260-227">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="2f260-227">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="2f260-228">String</span><span class="sxs-lookup"><span data-stu-id="2f260-228">String</span></span>|<span data-ttu-id="2f260-229">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="2f260-229">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="2f260-230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2f260-230">Requirements</span></span>

|<span data-ttu-id="2f260-231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2f260-231">Requirement</span></span>| <span data-ttu-id="2f260-232">Valeur</span><span class="sxs-lookup"><span data-stu-id="2f260-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="2f260-233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2f260-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2f260-234">1.1</span><span class="sxs-lookup"><span data-stu-id="2f260-234">1.1</span></span>|
|[<span data-ttu-id="2f260-235">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2f260-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2f260-236">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2f260-236">Compose or Read</span></span>|
