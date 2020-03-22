---
title: Espace de noms Office-ensemble de conditions requises 1,7
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,7.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 7991fd56097bbdebbfd4d4494a900626a1d3e02b
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891249"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="f8d97-103">Office (boîte aux lettres requise définie sur 1,7)</span><span class="sxs-lookup"><span data-stu-id="f8d97-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="f8d97-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f8d97-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f8d97-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f8d97-106">Requirements</span></span>

|<span data-ttu-id="f8d97-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8d97-107">Requirement</span></span>| <span data-ttu-id="f8d97-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8d97-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8d97-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8d97-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f8d97-110">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-110">1.1</span></span>|
|[<span data-ttu-id="f8d97-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8d97-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f8d97-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f8d97-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="f8d97-113">Properties</span></span>

| <span data-ttu-id="f8d97-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="f8d97-114">Property</span></span> | <span data-ttu-id="f8d97-115">Modes</span><span class="sxs-lookup"><span data-stu-id="f8d97-115">Modes</span></span> | <span data-ttu-id="f8d97-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f8d97-116">Return type</span></span> | <span data-ttu-id="f8d97-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="f8d97-117">Minimum</span></span><br><span data-ttu-id="f8d97-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8d97-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f8d97-119">context</span><span class="sxs-lookup"><span data-stu-id="f8d97-119">context</span></span>](office.context.md) | <span data-ttu-id="f8d97-120">Composition</span><span class="sxs-lookup"><span data-stu-id="f8d97-120">Compose</span></span><br><span data-ttu-id="f8d97-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-121">Read</span></span> | [<span data-ttu-id="f8d97-122">Context</span><span class="sxs-lookup"><span data-stu-id="f8d97-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7) | [<span data-ttu-id="f8d97-123">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f8d97-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="f8d97-124">Enumerations</span></span>

| <span data-ttu-id="f8d97-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="f8d97-125">Enumeration</span></span> | <span data-ttu-id="f8d97-126">Modes</span><span class="sxs-lookup"><span data-stu-id="f8d97-126">Modes</span></span> | <span data-ttu-id="f8d97-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f8d97-127">Return type</span></span> | <span data-ttu-id="f8d97-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="f8d97-128">Minimum</span></span><br><span data-ttu-id="f8d97-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8d97-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f8d97-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f8d97-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f8d97-131">Composition</span><span class="sxs-lookup"><span data-stu-id="f8d97-131">Compose</span></span><br><span data-ttu-id="f8d97-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-132">Read</span></span> | <span data-ttu-id="f8d97-133">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-133">String</span></span> | [<span data-ttu-id="f8d97-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f8d97-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f8d97-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f8d97-136">Composition</span><span class="sxs-lookup"><span data-stu-id="f8d97-136">Compose</span></span><br><span data-ttu-id="f8d97-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-137">Read</span></span> | <span data-ttu-id="f8d97-138">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-138">String</span></span> | [<span data-ttu-id="f8d97-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f8d97-140">EventType</span><span class="sxs-lookup"><span data-stu-id="f8d97-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="f8d97-141">Composition</span><span class="sxs-lookup"><span data-stu-id="f8d97-141">Compose</span></span><br><span data-ttu-id="f8d97-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-142">Read</span></span> | <span data-ttu-id="f8d97-143">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-143">String</span></span> | [<span data-ttu-id="f8d97-144">1,5</span><span class="sxs-lookup"><span data-stu-id="f8d97-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="f8d97-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f8d97-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f8d97-146">Composition</span><span class="sxs-lookup"><span data-stu-id="f8d97-146">Compose</span></span><br><span data-ttu-id="f8d97-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-147">Read</span></span> | <span data-ttu-id="f8d97-148">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-148">String</span></span> | [<span data-ttu-id="f8d97-149">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f8d97-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="f8d97-150">Namespaces</span></span>

<span data-ttu-id="f8d97-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f8d97-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f8d97-152">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="f8d97-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f8d97-153">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="f8d97-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="f8d97-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f8d97-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f8d97-155">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-155">Type</span></span>

*   <span data-ttu-id="f8d97-156">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f8d97-157">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f8d97-157">Properties:</span></span>

|<span data-ttu-id="f8d97-158">Nom</span><span class="sxs-lookup"><span data-stu-id="f8d97-158">Name</span></span>| <span data-ttu-id="f8d97-159">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-159">Type</span></span>| <span data-ttu-id="f8d97-160">Description</span><span class="sxs-lookup"><span data-stu-id="f8d97-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f8d97-161">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-161">String</span></span>|<span data-ttu-id="f8d97-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="f8d97-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f8d97-163">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-163">String</span></span>|<span data-ttu-id="f8d97-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="f8d97-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8d97-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f8d97-165">Requirements</span></span>

|<span data-ttu-id="f8d97-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8d97-166">Requirement</span></span>| <span data-ttu-id="f8d97-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8d97-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8d97-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8d97-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f8d97-169">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-169">1.1</span></span>|
|[<span data-ttu-id="f8d97-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8d97-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f8d97-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f8d97-172">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="f8d97-172">CoercionType: String</span></span>

<span data-ttu-id="f8d97-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f8d97-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f8d97-174">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-174">Type</span></span>

*   <span data-ttu-id="f8d97-175">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f8d97-176">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f8d97-176">Properties:</span></span>

|<span data-ttu-id="f8d97-177">Nom</span><span class="sxs-lookup"><span data-stu-id="f8d97-177">Name</span></span>| <span data-ttu-id="f8d97-178">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-178">Type</span></span>| <span data-ttu-id="f8d97-179">Description</span><span class="sxs-lookup"><span data-stu-id="f8d97-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f8d97-180">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-180">String</span></span>|<span data-ttu-id="f8d97-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f8d97-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f8d97-182">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-182">String</span></span>|<span data-ttu-id="f8d97-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="f8d97-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8d97-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f8d97-184">Requirements</span></span>

|<span data-ttu-id="f8d97-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8d97-185">Requirement</span></span>| <span data-ttu-id="f8d97-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8d97-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8d97-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8d97-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f8d97-188">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-188">1.1</span></span>|
|[<span data-ttu-id="f8d97-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8d97-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f8d97-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="f8d97-191">EventType : chaîne</span><span class="sxs-lookup"><span data-stu-id="f8d97-191">EventType: String</span></span>

<span data-ttu-id="f8d97-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="f8d97-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="f8d97-193">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-193">Type</span></span>

*   <span data-ttu-id="f8d97-194">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f8d97-195">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f8d97-195">Properties:</span></span>

| <span data-ttu-id="f8d97-196">Nom</span><span class="sxs-lookup"><span data-stu-id="f8d97-196">Name</span></span> | <span data-ttu-id="f8d97-197">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-197">Type</span></span> | <span data-ttu-id="f8d97-198">Description</span><span class="sxs-lookup"><span data-stu-id="f8d97-198">Description</span></span> | <span data-ttu-id="f8d97-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="f8d97-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="f8d97-200">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-200">String</span></span> | <span data-ttu-id="f8d97-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="f8d97-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="f8d97-202">1.7</span><span class="sxs-lookup"><span data-stu-id="f8d97-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="f8d97-203">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-203">String</span></span> | <span data-ttu-id="f8d97-204">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="f8d97-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="f8d97-205">1,5</span><span class="sxs-lookup"><span data-stu-id="f8d97-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="f8d97-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f8d97-206">String</span></span> | <span data-ttu-id="f8d97-207">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="f8d97-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="f8d97-208">1.7</span><span class="sxs-lookup"><span data-stu-id="f8d97-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="f8d97-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="f8d97-209">String</span></span> | <span data-ttu-id="f8d97-210">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="f8d97-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="f8d97-211">1.7</span><span class="sxs-lookup"><span data-stu-id="f8d97-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f8d97-212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f8d97-212">Requirements</span></span>

|<span data-ttu-id="f8d97-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8d97-213">Requirement</span></span>| <span data-ttu-id="f8d97-214">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8d97-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8d97-215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8d97-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f8d97-216">1,5</span><span class="sxs-lookup"><span data-stu-id="f8d97-216">1.5</span></span> |
|[<span data-ttu-id="f8d97-217">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8d97-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f8d97-218">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f8d97-219">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="f8d97-219">SourceProperty: String</span></span>

<span data-ttu-id="f8d97-220">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f8d97-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f8d97-221">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-221">Type</span></span>

*   <span data-ttu-id="f8d97-222">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f8d97-223">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f8d97-223">Properties:</span></span>

|<span data-ttu-id="f8d97-224">Nom</span><span class="sxs-lookup"><span data-stu-id="f8d97-224">Name</span></span>| <span data-ttu-id="f8d97-225">Type</span><span class="sxs-lookup"><span data-stu-id="f8d97-225">Type</span></span>| <span data-ttu-id="f8d97-226">Description</span><span class="sxs-lookup"><span data-stu-id="f8d97-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f8d97-227">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-227">String</span></span>|<span data-ttu-id="f8d97-228">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f8d97-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f8d97-229">String</span><span class="sxs-lookup"><span data-stu-id="f8d97-229">String</span></span>|<span data-ttu-id="f8d97-230">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="f8d97-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f8d97-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f8d97-231">Requirements</span></span>

|<span data-ttu-id="f8d97-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f8d97-232">Requirement</span></span>| <span data-ttu-id="f8d97-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="f8d97-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="f8d97-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f8d97-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f8d97-235">1.1</span><span class="sxs-lookup"><span data-stu-id="f8d97-235">1.1</span></span>|
|[<span data-ttu-id="f8d97-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f8d97-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f8d97-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f8d97-237">Compose or Read</span></span>|
