---
title: Office de noms - ensemble de conditions requises 1.8
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.8.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 00e236bed7e00159be8c94f727ca64ccaecd07b0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590525"
---
# <a name="office-mailbox-requirement-set-18"></a><span data-ttu-id="9a184-103">Office (ensemble de conditions requises de boîte aux lettres 1.8)</span><span class="sxs-lookup"><span data-stu-id="9a184-103">Office (Mailbox requirement set 1.8)</span></span>

<span data-ttu-id="9a184-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9a184-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a184-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a184-106">Requirements</span></span>

|<span data-ttu-id="9a184-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a184-107">Requirement</span></span>| <span data-ttu-id="9a184-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a184-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a184-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a184-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a184-110">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-110">1.1</span></span>|
|[<span data-ttu-id="9a184-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a184-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a184-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a184-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="9a184-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="9a184-113">Properties</span></span>

| <span data-ttu-id="9a184-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="9a184-114">Property</span></span> | <span data-ttu-id="9a184-115">Modes</span><span class="sxs-lookup"><span data-stu-id="9a184-115">Modes</span></span> | <span data-ttu-id="9a184-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="9a184-116">Return type</span></span> | <span data-ttu-id="9a184-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="9a184-117">Minimum</span></span><br><span data-ttu-id="9a184-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a184-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9a184-119">context</span><span class="sxs-lookup"><span data-stu-id="9a184-119">context</span></span>](office.context.md) | <span data-ttu-id="9a184-120">Composition</span><span class="sxs-lookup"><span data-stu-id="9a184-120">Compose</span></span><br><span data-ttu-id="9a184-121">Lire</span><span class="sxs-lookup"><span data-stu-id="9a184-121">Read</span></span> | [<span data-ttu-id="9a184-122">Context</span><span class="sxs-lookup"><span data-stu-id="9a184-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="9a184-123">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="9a184-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="9a184-124">Enumerations</span></span>

| <span data-ttu-id="9a184-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="9a184-125">Enumeration</span></span> | <span data-ttu-id="9a184-126">Modes</span><span class="sxs-lookup"><span data-stu-id="9a184-126">Modes</span></span> | <span data-ttu-id="9a184-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="9a184-127">Return type</span></span> | <span data-ttu-id="9a184-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="9a184-128">Minimum</span></span><br><span data-ttu-id="9a184-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a184-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9a184-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9a184-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9a184-131">Composition</span><span class="sxs-lookup"><span data-stu-id="9a184-131">Compose</span></span><br><span data-ttu-id="9a184-132">Lire</span><span class="sxs-lookup"><span data-stu-id="9a184-132">Read</span></span> | <span data-ttu-id="9a184-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-133">String</span></span> | [<span data-ttu-id="9a184-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a184-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9a184-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9a184-136">Composition</span><span class="sxs-lookup"><span data-stu-id="9a184-136">Compose</span></span><br><span data-ttu-id="9a184-137">Lire</span><span class="sxs-lookup"><span data-stu-id="9a184-137">Read</span></span> | <span data-ttu-id="9a184-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-138">String</span></span> | [<span data-ttu-id="9a184-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a184-140">EventType</span><span class="sxs-lookup"><span data-stu-id="9a184-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9a184-141">Composition</span><span class="sxs-lookup"><span data-stu-id="9a184-141">Compose</span></span><br><span data-ttu-id="9a184-142">Lire</span><span class="sxs-lookup"><span data-stu-id="9a184-142">Read</span></span> | <span data-ttu-id="9a184-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-143">String</span></span> | [<span data-ttu-id="9a184-144">1.5</span><span class="sxs-lookup"><span data-stu-id="9a184-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9a184-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9a184-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9a184-146">Composition</span><span class="sxs-lookup"><span data-stu-id="9a184-146">Compose</span></span><br><span data-ttu-id="9a184-147">Lire</span><span class="sxs-lookup"><span data-stu-id="9a184-147">Read</span></span> | <span data-ttu-id="9a184-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-148">String</span></span> | [<span data-ttu-id="9a184-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="9a184-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="9a184-150">Namespaces</span></span>

<span data-ttu-id="9a184-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="9a184-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="9a184-152">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="9a184-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="9a184-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="9a184-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="9a184-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="9a184-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9a184-155">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-155">Type</span></span>

*   <span data-ttu-id="9a184-156">String</span><span class="sxs-lookup"><span data-stu-id="9a184-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9a184-157">Propriétés</span><span class="sxs-lookup"><span data-stu-id="9a184-157">Properties</span></span>

|<span data-ttu-id="9a184-158">Nom</span><span class="sxs-lookup"><span data-stu-id="9a184-158">Name</span></span>| <span data-ttu-id="9a184-159">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-159">Type</span></span>| <span data-ttu-id="9a184-160">Description</span><span class="sxs-lookup"><span data-stu-id="9a184-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9a184-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-161">String</span></span>|<span data-ttu-id="9a184-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="9a184-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9a184-163">String</span><span class="sxs-lookup"><span data-stu-id="9a184-163">String</span></span>|<span data-ttu-id="9a184-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="9a184-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a184-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a184-165">Requirements</span></span>

|<span data-ttu-id="9a184-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a184-166">Requirement</span></span>| <span data-ttu-id="9a184-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a184-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a184-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a184-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a184-169">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-169">1.1</span></span>|
|[<span data-ttu-id="9a184-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a184-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a184-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a184-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="9a184-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="9a184-172">CoercionType: String</span></span>

<span data-ttu-id="9a184-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="9a184-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9a184-174">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-174">Type</span></span>

*   <span data-ttu-id="9a184-175">String</span><span class="sxs-lookup"><span data-stu-id="9a184-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9a184-176">Propriétés</span><span class="sxs-lookup"><span data-stu-id="9a184-176">Properties</span></span>

|<span data-ttu-id="9a184-177">Nom</span><span class="sxs-lookup"><span data-stu-id="9a184-177">Name</span></span>| <span data-ttu-id="9a184-178">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-178">Type</span></span>| <span data-ttu-id="9a184-179">Description</span><span class="sxs-lookup"><span data-stu-id="9a184-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9a184-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-180">String</span></span>|<span data-ttu-id="9a184-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="9a184-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9a184-182">String</span><span class="sxs-lookup"><span data-stu-id="9a184-182">String</span></span>|<span data-ttu-id="9a184-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="9a184-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a184-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a184-184">Requirements</span></span>

|<span data-ttu-id="9a184-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a184-185">Requirement</span></span>| <span data-ttu-id="9a184-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a184-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a184-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a184-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a184-188">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-188">1.1</span></span>|
|[<span data-ttu-id="9a184-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a184-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a184-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a184-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="9a184-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="9a184-191">EventType: String</span></span>

<span data-ttu-id="9a184-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="9a184-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9a184-193">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-193">Type</span></span>

*   <span data-ttu-id="9a184-194">String</span><span class="sxs-lookup"><span data-stu-id="9a184-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9a184-195">Propriétés</span><span class="sxs-lookup"><span data-stu-id="9a184-195">Properties</span></span>

| <span data-ttu-id="9a184-196">Nom</span><span class="sxs-lookup"><span data-stu-id="9a184-196">Name</span></span> | <span data-ttu-id="9a184-197">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-197">Type</span></span> | <span data-ttu-id="9a184-198">Description</span><span class="sxs-lookup"><span data-stu-id="9a184-198">Description</span></span> | <span data-ttu-id="9a184-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="9a184-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="9a184-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-200">String</span></span> | <span data-ttu-id="9a184-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="9a184-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="9a184-202">1.7</span><span class="sxs-lookup"><span data-stu-id="9a184-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="9a184-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-203">String</span></span> | <span data-ttu-id="9a184-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="9a184-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="9a184-205">1.8</span><span class="sxs-lookup"><span data-stu-id="9a184-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="9a184-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-206">String</span></span> | <span data-ttu-id="9a184-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="9a184-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="9a184-208">1.8</span><span class="sxs-lookup"><span data-stu-id="9a184-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="9a184-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-209">String</span></span> | <span data-ttu-id="9a184-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="9a184-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="9a184-211">1,5</span><span class="sxs-lookup"><span data-stu-id="9a184-211">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="9a184-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-212">String</span></span> | <span data-ttu-id="9a184-213">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="9a184-213">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="9a184-214">1.7</span><span class="sxs-lookup"><span data-stu-id="9a184-214">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="9a184-215">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-215">String</span></span> | <span data-ttu-id="9a184-216">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="9a184-216">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="9a184-217">1.7</span><span class="sxs-lookup"><span data-stu-id="9a184-217">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9a184-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a184-218">Requirements</span></span>

|<span data-ttu-id="9a184-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a184-219">Requirement</span></span>| <span data-ttu-id="9a184-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a184-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a184-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a184-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a184-222">1,5</span><span class="sxs-lookup"><span data-stu-id="9a184-222">1.5</span></span> |
|[<span data-ttu-id="9a184-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a184-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a184-224">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a184-224">Compose or Read</span></span> |

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="9a184-225">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="9a184-225">SourceProperty: String</span></span>

<span data-ttu-id="9a184-226">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="9a184-226">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9a184-227">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-227">Type</span></span>

*   <span data-ttu-id="9a184-228">String</span><span class="sxs-lookup"><span data-stu-id="9a184-228">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9a184-229">Propriétés</span><span class="sxs-lookup"><span data-stu-id="9a184-229">Properties</span></span>

|<span data-ttu-id="9a184-230">Nom</span><span class="sxs-lookup"><span data-stu-id="9a184-230">Name</span></span>| <span data-ttu-id="9a184-231">Type</span><span class="sxs-lookup"><span data-stu-id="9a184-231">Type</span></span>| <span data-ttu-id="9a184-232">Description</span><span class="sxs-lookup"><span data-stu-id="9a184-232">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9a184-233">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9a184-233">String</span></span>|<span data-ttu-id="9a184-234">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="9a184-234">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9a184-235">String</span><span class="sxs-lookup"><span data-stu-id="9a184-235">String</span></span>|<span data-ttu-id="9a184-236">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="9a184-236">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9a184-237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9a184-237">Requirements</span></span>

|<span data-ttu-id="9a184-238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9a184-238">Requirement</span></span>| <span data-ttu-id="9a184-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="9a184-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a184-240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9a184-240">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a184-241">1.1</span><span class="sxs-lookup"><span data-stu-id="9a184-241">1.1</span></span>|
|[<span data-ttu-id="9a184-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9a184-242">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a184-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9a184-243">Compose or Read</span></span>|
