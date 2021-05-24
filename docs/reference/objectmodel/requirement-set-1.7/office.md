---
title: Office de noms - ensemble de conditions requises 1.7
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.7.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 19c80c0c8c4aaf31c42aad16b3f474e92b7cdaec
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590973"
---
# <a name="office-mailbox-requirement-set-17"></a><span data-ttu-id="41989-103">Office (ensemble de conditions requises de boîte aux lettres 1.7)</span><span class="sxs-lookup"><span data-stu-id="41989-103">Office (Mailbox requirement set 1.7)</span></span>

<span data-ttu-id="41989-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="41989-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="41989-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="41989-106">Requirements</span></span>

|<span data-ttu-id="41989-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="41989-107">Requirement</span></span>| <span data-ttu-id="41989-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="41989-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="41989-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="41989-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41989-110">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-110">1.1</span></span>|
|[<span data-ttu-id="41989-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="41989-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41989-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="41989-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="41989-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="41989-113">Properties</span></span>

| <span data-ttu-id="41989-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="41989-114">Property</span></span> | <span data-ttu-id="41989-115">Modes</span><span class="sxs-lookup"><span data-stu-id="41989-115">Modes</span></span> | <span data-ttu-id="41989-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="41989-116">Return type</span></span> | <span data-ttu-id="41989-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="41989-117">Minimum</span></span><br><span data-ttu-id="41989-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="41989-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="41989-119">context</span><span class="sxs-lookup"><span data-stu-id="41989-119">context</span></span>](office.context.md) | <span data-ttu-id="41989-120">Composition</span><span class="sxs-lookup"><span data-stu-id="41989-120">Compose</span></span><br><span data-ttu-id="41989-121">Lire</span><span class="sxs-lookup"><span data-stu-id="41989-121">Read</span></span> | [<span data-ttu-id="41989-122">Context</span><span class="sxs-lookup"><span data-stu-id="41989-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="41989-123">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="41989-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="41989-124">Enumerations</span></span>

| <span data-ttu-id="41989-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="41989-125">Enumeration</span></span> | <span data-ttu-id="41989-126">Modes</span><span class="sxs-lookup"><span data-stu-id="41989-126">Modes</span></span> | <span data-ttu-id="41989-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="41989-127">Return type</span></span> | <span data-ttu-id="41989-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="41989-128">Minimum</span></span><br><span data-ttu-id="41989-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="41989-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="41989-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="41989-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="41989-131">Composition</span><span class="sxs-lookup"><span data-stu-id="41989-131">Compose</span></span><br><span data-ttu-id="41989-132">Lire</span><span class="sxs-lookup"><span data-stu-id="41989-132">Read</span></span> | <span data-ttu-id="41989-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-133">String</span></span> | [<span data-ttu-id="41989-134">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="41989-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="41989-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="41989-136">Composition</span><span class="sxs-lookup"><span data-stu-id="41989-136">Compose</span></span><br><span data-ttu-id="41989-137">Lire</span><span class="sxs-lookup"><span data-stu-id="41989-137">Read</span></span> | <span data-ttu-id="41989-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-138">String</span></span> | [<span data-ttu-id="41989-139">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="41989-140">EventType</span><span class="sxs-lookup"><span data-stu-id="41989-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="41989-141">Composition</span><span class="sxs-lookup"><span data-stu-id="41989-141">Compose</span></span><br><span data-ttu-id="41989-142">Lire</span><span class="sxs-lookup"><span data-stu-id="41989-142">Read</span></span> | <span data-ttu-id="41989-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-143">String</span></span> | [<span data-ttu-id="41989-144">1.5</span><span class="sxs-lookup"><span data-stu-id="41989-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="41989-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="41989-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="41989-146">Composition</span><span class="sxs-lookup"><span data-stu-id="41989-146">Compose</span></span><br><span data-ttu-id="41989-147">Lire</span><span class="sxs-lookup"><span data-stu-id="41989-147">Read</span></span> | <span data-ttu-id="41989-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-148">String</span></span> | [<span data-ttu-id="41989-149">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="41989-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="41989-150">Namespaces</span></span>

<span data-ttu-id="41989-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="41989-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="41989-152">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="41989-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="41989-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="41989-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="41989-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="41989-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="41989-155">Type</span><span class="sxs-lookup"><span data-stu-id="41989-155">Type</span></span>

*   <span data-ttu-id="41989-156">String</span><span class="sxs-lookup"><span data-stu-id="41989-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41989-157">Propriétés</span><span class="sxs-lookup"><span data-stu-id="41989-157">Properties</span></span>

|<span data-ttu-id="41989-158">Nom</span><span class="sxs-lookup"><span data-stu-id="41989-158">Name</span></span>| <span data-ttu-id="41989-159">Type</span><span class="sxs-lookup"><span data-stu-id="41989-159">Type</span></span>| <span data-ttu-id="41989-160">Description</span><span class="sxs-lookup"><span data-stu-id="41989-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="41989-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-161">String</span></span>|<span data-ttu-id="41989-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="41989-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="41989-163">String</span><span class="sxs-lookup"><span data-stu-id="41989-163">String</span></span>|<span data-ttu-id="41989-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="41989-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41989-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="41989-165">Requirements</span></span>

|<span data-ttu-id="41989-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="41989-166">Requirement</span></span>| <span data-ttu-id="41989-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="41989-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="41989-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="41989-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41989-169">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-169">1.1</span></span>|
|[<span data-ttu-id="41989-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="41989-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41989-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="41989-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="41989-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="41989-172">CoercionType: String</span></span>

<span data-ttu-id="41989-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="41989-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="41989-174">Type</span><span class="sxs-lookup"><span data-stu-id="41989-174">Type</span></span>

*   <span data-ttu-id="41989-175">String</span><span class="sxs-lookup"><span data-stu-id="41989-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41989-176">Propriétés</span><span class="sxs-lookup"><span data-stu-id="41989-176">Properties</span></span>

|<span data-ttu-id="41989-177">Nom</span><span class="sxs-lookup"><span data-stu-id="41989-177">Name</span></span>| <span data-ttu-id="41989-178">Type</span><span class="sxs-lookup"><span data-stu-id="41989-178">Type</span></span>| <span data-ttu-id="41989-179">Description</span><span class="sxs-lookup"><span data-stu-id="41989-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="41989-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-180">String</span></span>|<span data-ttu-id="41989-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="41989-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="41989-182">String</span><span class="sxs-lookup"><span data-stu-id="41989-182">String</span></span>|<span data-ttu-id="41989-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="41989-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41989-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="41989-184">Requirements</span></span>

|<span data-ttu-id="41989-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="41989-185">Requirement</span></span>| <span data-ttu-id="41989-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="41989-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="41989-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="41989-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41989-188">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-188">1.1</span></span>|
|[<span data-ttu-id="41989-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="41989-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41989-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="41989-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="41989-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="41989-191">EventType: String</span></span>

<span data-ttu-id="41989-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="41989-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="41989-193">Type</span><span class="sxs-lookup"><span data-stu-id="41989-193">Type</span></span>

*   <span data-ttu-id="41989-194">String</span><span class="sxs-lookup"><span data-stu-id="41989-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41989-195">Propriétés</span><span class="sxs-lookup"><span data-stu-id="41989-195">Properties</span></span>

| <span data-ttu-id="41989-196">Nom</span><span class="sxs-lookup"><span data-stu-id="41989-196">Name</span></span> | <span data-ttu-id="41989-197">Type</span><span class="sxs-lookup"><span data-stu-id="41989-197">Type</span></span> | <span data-ttu-id="41989-198">Description</span><span class="sxs-lookup"><span data-stu-id="41989-198">Description</span></span> | <span data-ttu-id="41989-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="41989-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="41989-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-200">String</span></span> | <span data-ttu-id="41989-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="41989-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="41989-202">1.7</span><span class="sxs-lookup"><span data-stu-id="41989-202">1.7</span></span> |
|`ItemChanged`| <span data-ttu-id="41989-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-203">String</span></span> | <span data-ttu-id="41989-204">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="41989-204">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="41989-205">1,5</span><span class="sxs-lookup"><span data-stu-id="41989-205">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="41989-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-206">String</span></span> | <span data-ttu-id="41989-207">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="41989-207">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="41989-208">1.7</span><span class="sxs-lookup"><span data-stu-id="41989-208">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="41989-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-209">String</span></span> | <span data-ttu-id="41989-210">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="41989-210">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="41989-211">1.7</span><span class="sxs-lookup"><span data-stu-id="41989-211">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="41989-212">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="41989-212">Requirements</span></span>

|<span data-ttu-id="41989-213">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="41989-213">Requirement</span></span>| <span data-ttu-id="41989-214">Valeur</span><span class="sxs-lookup"><span data-stu-id="41989-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="41989-215">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="41989-215">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41989-216">1,5</span><span class="sxs-lookup"><span data-stu-id="41989-216">1.5</span></span> |
|[<span data-ttu-id="41989-217">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="41989-217">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41989-218">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="41989-218">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="41989-219">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="41989-219">SourceProperty: String</span></span>

<span data-ttu-id="41989-220">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="41989-220">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="41989-221">Type</span><span class="sxs-lookup"><span data-stu-id="41989-221">Type</span></span>

*   <span data-ttu-id="41989-222">String</span><span class="sxs-lookup"><span data-stu-id="41989-222">String</span></span>

##### <a name="properties"></a><span data-ttu-id="41989-223">Propriétés</span><span class="sxs-lookup"><span data-stu-id="41989-223">Properties</span></span>

|<span data-ttu-id="41989-224">Nom</span><span class="sxs-lookup"><span data-stu-id="41989-224">Name</span></span>| <span data-ttu-id="41989-225">Type</span><span class="sxs-lookup"><span data-stu-id="41989-225">Type</span></span>| <span data-ttu-id="41989-226">Description</span><span class="sxs-lookup"><span data-stu-id="41989-226">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="41989-227">Chaîne</span><span class="sxs-lookup"><span data-stu-id="41989-227">String</span></span>|<span data-ttu-id="41989-228">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="41989-228">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="41989-229">String</span><span class="sxs-lookup"><span data-stu-id="41989-229">String</span></span>|<span data-ttu-id="41989-230">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="41989-230">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="41989-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="41989-231">Requirements</span></span>

|<span data-ttu-id="41989-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="41989-232">Requirement</span></span>| <span data-ttu-id="41989-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="41989-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="41989-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="41989-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="41989-235">1.1</span><span class="sxs-lookup"><span data-stu-id="41989-235">1.1</span></span>|
|[<span data-ttu-id="41989-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="41989-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="41989-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="41989-237">Compose or Read</span></span>|
