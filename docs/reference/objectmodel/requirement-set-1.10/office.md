---
title: Office de noms - ensemble de conditions requises 1.10
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.10.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: e7b7ab9127ebf8ce9b7394d348144fe63b47de6c
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592037"
---
# <a name="office-mailbox-requirement-set-110"></a><span data-ttu-id="15953-103">Office (ensemble de conditions requises de boîte aux lettres 1.10)</span><span class="sxs-lookup"><span data-stu-id="15953-103">Office (Mailbox requirement set 1.10)</span></span>

<span data-ttu-id="15953-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="15953-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="15953-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="15953-106">Requirements</span></span>

|<span data-ttu-id="15953-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="15953-107">Requirement</span></span>| <span data-ttu-id="15953-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="15953-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="15953-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="15953-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="15953-110">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-110">1.1</span></span>|
|[<span data-ttu-id="15953-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="15953-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="15953-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="15953-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="15953-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="15953-113">Properties</span></span>

| <span data-ttu-id="15953-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="15953-114">Property</span></span> | <span data-ttu-id="15953-115">Modes</span><span class="sxs-lookup"><span data-stu-id="15953-115">Modes</span></span> | <span data-ttu-id="15953-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="15953-116">Return type</span></span> | <span data-ttu-id="15953-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="15953-117">Minimum</span></span><br><span data-ttu-id="15953-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="15953-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="15953-119">context</span><span class="sxs-lookup"><span data-stu-id="15953-119">context</span></span>](office.context.md) | <span data-ttu-id="15953-120">Composition</span><span class="sxs-lookup"><span data-stu-id="15953-120">Compose</span></span><br><span data-ttu-id="15953-121">Lire</span><span class="sxs-lookup"><span data-stu-id="15953-121">Read</span></span> | [<span data-ttu-id="15953-122">Context</span><span class="sxs-lookup"><span data-stu-id="15953-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="15953-123">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="15953-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="15953-124">Enumerations</span></span>

| <span data-ttu-id="15953-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="15953-125">Enumeration</span></span> | <span data-ttu-id="15953-126">Modes</span><span class="sxs-lookup"><span data-stu-id="15953-126">Modes</span></span> | <span data-ttu-id="15953-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="15953-127">Return type</span></span> | <span data-ttu-id="15953-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="15953-128">Minimum</span></span><br><span data-ttu-id="15953-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="15953-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="15953-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="15953-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="15953-131">Composition</span><span class="sxs-lookup"><span data-stu-id="15953-131">Compose</span></span><br><span data-ttu-id="15953-132">Lire</span><span class="sxs-lookup"><span data-stu-id="15953-132">Read</span></span> | <span data-ttu-id="15953-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-133">String</span></span> | [<span data-ttu-id="15953-134">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="15953-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="15953-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="15953-136">Composition</span><span class="sxs-lookup"><span data-stu-id="15953-136">Compose</span></span><br><span data-ttu-id="15953-137">Lire</span><span class="sxs-lookup"><span data-stu-id="15953-137">Read</span></span> | <span data-ttu-id="15953-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-138">String</span></span> | [<span data-ttu-id="15953-139">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="15953-140">EventType</span><span class="sxs-lookup"><span data-stu-id="15953-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="15953-141">Composition</span><span class="sxs-lookup"><span data-stu-id="15953-141">Compose</span></span><br><span data-ttu-id="15953-142">Lire</span><span class="sxs-lookup"><span data-stu-id="15953-142">Read</span></span> | <span data-ttu-id="15953-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-143">String</span></span> | [<span data-ttu-id="15953-144">1.5</span><span class="sxs-lookup"><span data-stu-id="15953-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="15953-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="15953-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="15953-146">Composition</span><span class="sxs-lookup"><span data-stu-id="15953-146">Compose</span></span><br><span data-ttu-id="15953-147">Lire</span><span class="sxs-lookup"><span data-stu-id="15953-147">Read</span></span> | <span data-ttu-id="15953-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-148">String</span></span> | [<span data-ttu-id="15953-149">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="15953-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="15953-150">Namespaces</span></span>

<span data-ttu-id="15953-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="15953-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.10&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="15953-152">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="15953-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="15953-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="15953-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="15953-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="15953-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="15953-155">Type</span><span class="sxs-lookup"><span data-stu-id="15953-155">Type</span></span>

*   <span data-ttu-id="15953-156">String</span><span class="sxs-lookup"><span data-stu-id="15953-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="15953-157">Propriétés</span><span class="sxs-lookup"><span data-stu-id="15953-157">Properties</span></span>

|<span data-ttu-id="15953-158">Nom</span><span class="sxs-lookup"><span data-stu-id="15953-158">Name</span></span>| <span data-ttu-id="15953-159">Type</span><span class="sxs-lookup"><span data-stu-id="15953-159">Type</span></span>| <span data-ttu-id="15953-160">Description</span><span class="sxs-lookup"><span data-stu-id="15953-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="15953-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-161">String</span></span>|<span data-ttu-id="15953-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="15953-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="15953-163">String</span><span class="sxs-lookup"><span data-stu-id="15953-163">String</span></span>|<span data-ttu-id="15953-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="15953-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="15953-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="15953-165">Requirements</span></span>

|<span data-ttu-id="15953-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="15953-166">Requirement</span></span>| <span data-ttu-id="15953-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="15953-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="15953-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="15953-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="15953-169">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-169">1.1</span></span>|
|[<span data-ttu-id="15953-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="15953-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="15953-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="15953-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="15953-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="15953-172">CoercionType: String</span></span>

<span data-ttu-id="15953-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="15953-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="15953-174">Type</span><span class="sxs-lookup"><span data-stu-id="15953-174">Type</span></span>

*   <span data-ttu-id="15953-175">String</span><span class="sxs-lookup"><span data-stu-id="15953-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="15953-176">Propriétés</span><span class="sxs-lookup"><span data-stu-id="15953-176">Properties</span></span>

|<span data-ttu-id="15953-177">Nom</span><span class="sxs-lookup"><span data-stu-id="15953-177">Name</span></span>| <span data-ttu-id="15953-178">Type</span><span class="sxs-lookup"><span data-stu-id="15953-178">Type</span></span>| <span data-ttu-id="15953-179">Description</span><span class="sxs-lookup"><span data-stu-id="15953-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="15953-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-180">String</span></span>|<span data-ttu-id="15953-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="15953-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="15953-182">String</span><span class="sxs-lookup"><span data-stu-id="15953-182">String</span></span>|<span data-ttu-id="15953-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="15953-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="15953-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="15953-184">Requirements</span></span>

|<span data-ttu-id="15953-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="15953-185">Requirement</span></span>| <span data-ttu-id="15953-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="15953-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="15953-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="15953-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="15953-188">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-188">1.1</span></span>|
|[<span data-ttu-id="15953-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="15953-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="15953-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="15953-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="15953-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="15953-191">EventType: String</span></span>

<span data-ttu-id="15953-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="15953-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="15953-193">Type</span><span class="sxs-lookup"><span data-stu-id="15953-193">Type</span></span>

*   <span data-ttu-id="15953-194">String</span><span class="sxs-lookup"><span data-stu-id="15953-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="15953-195">Propriétés</span><span class="sxs-lookup"><span data-stu-id="15953-195">Properties</span></span>

| <span data-ttu-id="15953-196">Nom</span><span class="sxs-lookup"><span data-stu-id="15953-196">Name</span></span> | <span data-ttu-id="15953-197">Type</span><span class="sxs-lookup"><span data-stu-id="15953-197">Type</span></span> | <span data-ttu-id="15953-198">Description</span><span class="sxs-lookup"><span data-stu-id="15953-198">Description</span></span> | <span data-ttu-id="15953-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="15953-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="15953-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-200">String</span></span> | <span data-ttu-id="15953-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="15953-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="15953-202">1.7</span><span class="sxs-lookup"><span data-stu-id="15953-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="15953-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-203">String</span></span> | <span data-ttu-id="15953-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="15953-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="15953-205">1.8</span><span class="sxs-lookup"><span data-stu-id="15953-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="15953-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-206">String</span></span> | <span data-ttu-id="15953-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="15953-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="15953-208">1.8</span><span class="sxs-lookup"><span data-stu-id="15953-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="15953-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-209">String</span></span> | <span data-ttu-id="15953-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="15953-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="15953-211">1,5</span><span class="sxs-lookup"><span data-stu-id="15953-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="15953-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-212">String</span></span> | <span data-ttu-id="15953-213">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="15953-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="15953-214">1.10</span><span class="sxs-lookup"><span data-stu-id="15953-214">1.10</span></span> |
|`RecipientsChanged`| <span data-ttu-id="15953-215">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-215">String</span></span> | <span data-ttu-id="15953-216">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="15953-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="15953-217">1.7</span><span class="sxs-lookup"><span data-stu-id="15953-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="15953-218">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-218">String</span></span> | <span data-ttu-id="15953-219">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="15953-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="15953-220">1.7</span><span class="sxs-lookup"><span data-stu-id="15953-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="15953-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="15953-221">Requirements</span></span>

|<span data-ttu-id="15953-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="15953-222">Requirement</span></span>| <span data-ttu-id="15953-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="15953-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="15953-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="15953-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="15953-225">1,5</span><span class="sxs-lookup"><span data-stu-id="15953-225">1.5</span></span> |
|[<span data-ttu-id="15953-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="15953-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="15953-227">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="15953-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="15953-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="15953-228">SourceProperty: String</span></span>

<span data-ttu-id="15953-229">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="15953-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="15953-230">Type</span><span class="sxs-lookup"><span data-stu-id="15953-230">Type</span></span>

*   <span data-ttu-id="15953-231">String</span><span class="sxs-lookup"><span data-stu-id="15953-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="15953-232">Propriétés</span><span class="sxs-lookup"><span data-stu-id="15953-232">Properties</span></span>

|<span data-ttu-id="15953-233">Nom</span><span class="sxs-lookup"><span data-stu-id="15953-233">Name</span></span>| <span data-ttu-id="15953-234">Type</span><span class="sxs-lookup"><span data-stu-id="15953-234">Type</span></span>| <span data-ttu-id="15953-235">Description</span><span class="sxs-lookup"><span data-stu-id="15953-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="15953-236">Chaîne</span><span class="sxs-lookup"><span data-stu-id="15953-236">String</span></span>|<span data-ttu-id="15953-237">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="15953-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="15953-238">String</span><span class="sxs-lookup"><span data-stu-id="15953-238">String</span></span>|<span data-ttu-id="15953-239">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="15953-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="15953-240">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="15953-240">Requirements</span></span>

|<span data-ttu-id="15953-241">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="15953-241">Requirement</span></span>| <span data-ttu-id="15953-242">Valeur</span><span class="sxs-lookup"><span data-stu-id="15953-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="15953-243">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="15953-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="15953-244">1.1</span><span class="sxs-lookup"><span data-stu-id="15953-244">1.1</span></span>|
|[<span data-ttu-id="15953-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="15953-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="15953-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="15953-246">Compose or Read</span></span>|
