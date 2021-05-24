---
title: Espace de noms Office – ensemble de conditions requises
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises d’aperçu de l’API de boîte aux lettres.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 72e2300dd50ff01e26417efaca92906049358fc0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590882"
---
# <a name="office-mailbox-preview-requirement-set"></a><span data-ttu-id="e12e3-103">Office (ensemble de conditions requises pour l’aperçu de boîte aux lettres)</span><span class="sxs-lookup"><span data-stu-id="e12e3-103">Office (Mailbox preview requirement set)</span></span>

<span data-ttu-id="e12e3-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="e12e3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e12e3-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e12e3-106">Requirements</span></span>

|<span data-ttu-id="e12e3-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e12e3-107">Requirement</span></span>| <span data-ttu-id="e12e3-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="e12e3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="e12e3-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e12e3-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e12e3-110">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-110">1.1</span></span>|
|[<span data-ttu-id="e12e3-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e12e3-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e12e3-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e12e3-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="e12e3-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e12e3-113">Properties</span></span>

| <span data-ttu-id="e12e3-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="e12e3-114">Property</span></span> | <span data-ttu-id="e12e3-115">Modes</span><span class="sxs-lookup"><span data-stu-id="e12e3-115">Modes</span></span> | <span data-ttu-id="e12e3-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="e12e3-116">Return type</span></span> | <span data-ttu-id="e12e3-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="e12e3-117">Minimum</span></span><br><span data-ttu-id="e12e3-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e12e3-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e12e3-119">context</span><span class="sxs-lookup"><span data-stu-id="e12e3-119">context</span></span>](office.context.md) | <span data-ttu-id="e12e3-120">Composition</span><span class="sxs-lookup"><span data-stu-id="e12e3-120">Compose</span></span><br><span data-ttu-id="e12e3-121">Lire</span><span class="sxs-lookup"><span data-stu-id="e12e3-121">Read</span></span> | [<span data-ttu-id="e12e3-122">Context</span><span class="sxs-lookup"><span data-stu-id="e12e3-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="e12e3-123">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="e12e3-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="e12e3-124">Enumerations</span></span>

| <span data-ttu-id="e12e3-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="e12e3-125">Enumeration</span></span> | <span data-ttu-id="e12e3-126">Modes</span><span class="sxs-lookup"><span data-stu-id="e12e3-126">Modes</span></span> | <span data-ttu-id="e12e3-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="e12e3-127">Return type</span></span> | <span data-ttu-id="e12e3-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="e12e3-128">Minimum</span></span><br><span data-ttu-id="e12e3-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e12e3-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e12e3-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="e12e3-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="e12e3-131">Composition</span><span class="sxs-lookup"><span data-stu-id="e12e3-131">Compose</span></span><br><span data-ttu-id="e12e3-132">Lire</span><span class="sxs-lookup"><span data-stu-id="e12e3-132">Read</span></span> | <span data-ttu-id="e12e3-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-133">String</span></span> | [<span data-ttu-id="e12e3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e12e3-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="e12e3-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="e12e3-136">Composition</span><span class="sxs-lookup"><span data-stu-id="e12e3-136">Compose</span></span><br><span data-ttu-id="e12e3-137">Lire</span><span class="sxs-lookup"><span data-stu-id="e12e3-137">Read</span></span> | <span data-ttu-id="e12e3-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-138">String</span></span> | [<span data-ttu-id="e12e3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e12e3-140">EventType</span><span class="sxs-lookup"><span data-stu-id="e12e3-140">EventType</span></span>](#eventtype-string) | <span data-ttu-id="e12e3-141">Composition</span><span class="sxs-lookup"><span data-stu-id="e12e3-141">Compose</span></span><br><span data-ttu-id="e12e3-142">Lire</span><span class="sxs-lookup"><span data-stu-id="e12e3-142">Read</span></span> | <span data-ttu-id="e12e3-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-143">String</span></span> | [<span data-ttu-id="e12e3-144">1.5</span><span class="sxs-lookup"><span data-stu-id="e12e3-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e12e3-145">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="e12e3-145">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="e12e3-146">Composition</span><span class="sxs-lookup"><span data-stu-id="e12e3-146">Compose</span></span><br><span data-ttu-id="e12e3-147">Lire</span><span class="sxs-lookup"><span data-stu-id="e12e3-147">Read</span></span> | <span data-ttu-id="e12e3-148">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-148">String</span></span> | [<span data-ttu-id="e12e3-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="e12e3-150">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="e12e3-150">Namespaces</span></span>

<span data-ttu-id="e12e3-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="e12e3-151">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-preview&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="e12e3-152">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="e12e3-152">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="e12e3-153">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="e12e3-153">AsyncResultStatus: String</span></span>

<span data-ttu-id="e12e3-154">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="e12e3-154">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="e12e3-155">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-155">Type</span></span>

*   <span data-ttu-id="e12e3-156">String</span><span class="sxs-lookup"><span data-stu-id="e12e3-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e12e3-157">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e12e3-157">Properties</span></span>

|<span data-ttu-id="e12e3-158">Nom</span><span class="sxs-lookup"><span data-stu-id="e12e3-158">Name</span></span>| <span data-ttu-id="e12e3-159">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-159">Type</span></span>| <span data-ttu-id="e12e3-160">Description</span><span class="sxs-lookup"><span data-stu-id="e12e3-160">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="e12e3-161">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-161">String</span></span>|<span data-ttu-id="e12e3-162">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="e12e3-162">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="e12e3-163">String</span><span class="sxs-lookup"><span data-stu-id="e12e3-163">String</span></span>|<span data-ttu-id="e12e3-164">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="e12e3-164">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e12e3-165">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e12e3-165">Requirements</span></span>

|<span data-ttu-id="e12e3-166">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e12e3-166">Requirement</span></span>| <span data-ttu-id="e12e3-167">Valeur</span><span class="sxs-lookup"><span data-stu-id="e12e3-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="e12e3-168">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e12e3-168">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e12e3-169">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-169">1.1</span></span>|
|[<span data-ttu-id="e12e3-170">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e12e3-170">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e12e3-171">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e12e3-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="e12e3-172">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="e12e3-172">CoercionType: String</span></span>

<span data-ttu-id="e12e3-173">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e12e3-173">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e12e3-174">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-174">Type</span></span>

*   <span data-ttu-id="e12e3-175">String</span><span class="sxs-lookup"><span data-stu-id="e12e3-175">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e12e3-176">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e12e3-176">Properties</span></span>

|<span data-ttu-id="e12e3-177">Nom</span><span class="sxs-lookup"><span data-stu-id="e12e3-177">Name</span></span>| <span data-ttu-id="e12e3-178">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-178">Type</span></span>| <span data-ttu-id="e12e3-179">Description</span><span class="sxs-lookup"><span data-stu-id="e12e3-179">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="e12e3-180">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-180">String</span></span>|<span data-ttu-id="e12e3-181">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="e12e3-181">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="e12e3-182">String</span><span class="sxs-lookup"><span data-stu-id="e12e3-182">String</span></span>|<span data-ttu-id="e12e3-183">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="e12e3-183">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e12e3-184">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e12e3-184">Requirements</span></span>

|<span data-ttu-id="e12e3-185">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e12e3-185">Requirement</span></span>| <span data-ttu-id="e12e3-186">Valeur</span><span class="sxs-lookup"><span data-stu-id="e12e3-186">Value</span></span>|
|---|---|
|[<span data-ttu-id="e12e3-187">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e12e3-187">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e12e3-188">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-188">1.1</span></span>|
|[<span data-ttu-id="e12e3-189">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e12e3-189">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e12e3-190">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e12e3-190">Compose or Read</span></span>|

<br>

---
---

#### <a name="eventtype-string"></a><span data-ttu-id="e12e3-191">EventType: String</span><span class="sxs-lookup"><span data-stu-id="e12e3-191">EventType: String</span></span>

<span data-ttu-id="e12e3-192">spécifie l’événement associé à un gestionnaire d’événements.</span><span class="sxs-lookup"><span data-stu-id="e12e3-192">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="e12e3-193">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-193">Type</span></span>

*   <span data-ttu-id="e12e3-194">String</span><span class="sxs-lookup"><span data-stu-id="e12e3-194">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e12e3-195">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e12e3-195">Properties</span></span>

| <span data-ttu-id="e12e3-196">Nom</span><span class="sxs-lookup"><span data-stu-id="e12e3-196">Name</span></span> | <span data-ttu-id="e12e3-197">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-197">Type</span></span> | <span data-ttu-id="e12e3-198">Description</span><span class="sxs-lookup"><span data-stu-id="e12e3-198">Description</span></span> | <span data-ttu-id="e12e3-199">Ensemble de conditions requises minimales</span><span class="sxs-lookup"><span data-stu-id="e12e3-199">Minimum requirement set</span></span> |
|---|---|---|:---:|
|`AppointmentTimeChanged`| <span data-ttu-id="e12e3-200">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-200">String</span></span> | <span data-ttu-id="e12e3-201">La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="e12e3-201">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="e12e3-202">1.7</span><span class="sxs-lookup"><span data-stu-id="e12e3-202">1.7</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="e12e3-203">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-203">String</span></span> | <span data-ttu-id="e12e3-204">Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci.</span><span class="sxs-lookup"><span data-stu-id="e12e3-204">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="e12e3-205">1.8</span><span class="sxs-lookup"><span data-stu-id="e12e3-205">1.8</span></span> |
|`EnhancedLocationsChanged`| <span data-ttu-id="e12e3-206">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-206">String</span></span> | <span data-ttu-id="e12e3-207">L’emplacement du rendez-vous sélectionné a changé.</span><span class="sxs-lookup"><span data-stu-id="e12e3-207">The location of the selected appointment has changed.</span></span> | <span data-ttu-id="e12e3-208">1.8</span><span class="sxs-lookup"><span data-stu-id="e12e3-208">1.8</span></span> |
|`ItemChanged`| <span data-ttu-id="e12e3-209">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-209">String</span></span> | <span data-ttu-id="e12e3-210">Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé.</span><span class="sxs-lookup"><span data-stu-id="e12e3-210">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="e12e3-211">1,5</span><span class="sxs-lookup"><span data-stu-id="e12e3-211">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="e12e3-212">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-212">String</span></span> | <span data-ttu-id="e12e3-213">Le thème Office de la boîte aux lettres a été modifié.</span><span class="sxs-lookup"><span data-stu-id="e12e3-213">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="e12e3-214">Aperçu</span><span class="sxs-lookup"><span data-stu-id="e12e3-214">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="e12e3-215">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-215">String</span></span> | <span data-ttu-id="e12e3-216">La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié.</span><span class="sxs-lookup"><span data-stu-id="e12e3-216">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="e12e3-217">1.7</span><span class="sxs-lookup"><span data-stu-id="e12e3-217">1.7</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="e12e3-218">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-218">String</span></span> | <span data-ttu-id="e12e3-219">La périodicité de la série sélectionnée a été modifiée.</span><span class="sxs-lookup"><span data-stu-id="e12e3-219">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="e12e3-220">1.7</span><span class="sxs-lookup"><span data-stu-id="e12e3-220">1.7</span></span> |

##### <a name="requirements"></a><span data-ttu-id="e12e3-221">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e12e3-221">Requirements</span></span>

|<span data-ttu-id="e12e3-222">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e12e3-222">Requirement</span></span>| <span data-ttu-id="e12e3-223">Valeur</span><span class="sxs-lookup"><span data-stu-id="e12e3-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="e12e3-224">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e12e3-224">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e12e3-225">1,5</span><span class="sxs-lookup"><span data-stu-id="e12e3-225">1.5</span></span> |
|[<span data-ttu-id="e12e3-226">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e12e3-226">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e12e3-227">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e12e3-227">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="e12e3-228">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="e12e3-228">SourceProperty: String</span></span>

<span data-ttu-id="e12e3-229">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="e12e3-229">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="e12e3-230">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-230">Type</span></span>

*   <span data-ttu-id="e12e3-231">String</span><span class="sxs-lookup"><span data-stu-id="e12e3-231">String</span></span>

##### <a name="properties"></a><span data-ttu-id="e12e3-232">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e12e3-232">Properties</span></span>

|<span data-ttu-id="e12e3-233">Nom</span><span class="sxs-lookup"><span data-stu-id="e12e3-233">Name</span></span>| <span data-ttu-id="e12e3-234">Type</span><span class="sxs-lookup"><span data-stu-id="e12e3-234">Type</span></span>| <span data-ttu-id="e12e3-235">Description</span><span class="sxs-lookup"><span data-stu-id="e12e3-235">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="e12e3-236">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e12e3-236">String</span></span>|<span data-ttu-id="e12e3-237">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="e12e3-237">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="e12e3-238">String</span><span class="sxs-lookup"><span data-stu-id="e12e3-238">String</span></span>|<span data-ttu-id="e12e3-239">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="e12e3-239">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="e12e3-240">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e12e3-240">Requirements</span></span>

|<span data-ttu-id="e12e3-241">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e12e3-241">Requirement</span></span>| <span data-ttu-id="e12e3-242">Valeur</span><span class="sxs-lookup"><span data-stu-id="e12e3-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="e12e3-243">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e12e3-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e12e3-244">1.1</span><span class="sxs-lookup"><span data-stu-id="e12e3-244">1.1</span></span>|
|[<span data-ttu-id="e12e3-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e12e3-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e12e3-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e12e3-246">Compose or Read</span></span>|
