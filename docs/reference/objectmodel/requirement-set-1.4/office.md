---
title: Espace de noms Office-ensemble de conditions requises 1,4
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,4.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 8b0447b819a7360be195de1262c88877efc2f2fc
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891389"
---
# <a name="office-mailbox-requirement-set-14"></a><span data-ttu-id="f1c85-103">Office (boîte aux lettres requise définie sur 1,4)</span><span class="sxs-lookup"><span data-stu-id="f1c85-103">Office (Mailbox requirement set 1.4)</span></span>

<span data-ttu-id="f1c85-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f1c85-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f1c85-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f1c85-106">Requirements</span></span>

|<span data-ttu-id="f1c85-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f1c85-107">Requirement</span></span>| <span data-ttu-id="f1c85-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="f1c85-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1c85-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f1c85-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f1c85-110">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-110">1.1</span></span>|
|[<span data-ttu-id="f1c85-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f1c85-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f1c85-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f1c85-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="f1c85-113">Properties</span></span>

| <span data-ttu-id="f1c85-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="f1c85-114">Property</span></span> | <span data-ttu-id="f1c85-115">Modes</span><span class="sxs-lookup"><span data-stu-id="f1c85-115">Modes</span></span> | <span data-ttu-id="f1c85-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f1c85-116">Return type</span></span> | <span data-ttu-id="f1c85-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="f1c85-117">Minimum</span></span><br><span data-ttu-id="f1c85-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f1c85-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f1c85-119">context</span><span class="sxs-lookup"><span data-stu-id="f1c85-119">context</span></span>](office.context.md) | <span data-ttu-id="f1c85-120">Composition</span><span class="sxs-lookup"><span data-stu-id="f1c85-120">Compose</span></span><br><span data-ttu-id="f1c85-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-121">Read</span></span> | [<span data-ttu-id="f1c85-122">Context</span><span class="sxs-lookup"><span data-stu-id="f1c85-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.4) | [<span data-ttu-id="f1c85-123">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a><span data-ttu-id="f1c85-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="f1c85-124">Enumerations</span></span>

| <span data-ttu-id="f1c85-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="f1c85-125">Enumeration</span></span> | <span data-ttu-id="f1c85-126">Modes</span><span class="sxs-lookup"><span data-stu-id="f1c85-126">Modes</span></span> | <span data-ttu-id="f1c85-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="f1c85-127">Return type</span></span> | <span data-ttu-id="f1c85-128">Minimale</span><span class="sxs-lookup"><span data-stu-id="f1c85-128">Minimum</span></span><br><span data-ttu-id="f1c85-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="f1c85-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f1c85-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="f1c85-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="f1c85-131">Composition</span><span class="sxs-lookup"><span data-stu-id="f1c85-131">Compose</span></span><br><span data-ttu-id="f1c85-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-132">Read</span></span> | <span data-ttu-id="f1c85-133">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-133">String</span></span> | [<span data-ttu-id="f1c85-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f1c85-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="f1c85-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="f1c85-136">Composition</span><span class="sxs-lookup"><span data-stu-id="f1c85-136">Compose</span></span><br><span data-ttu-id="f1c85-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-137">Read</span></span> | <span data-ttu-id="f1c85-138">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-138">String</span></span> | [<span data-ttu-id="f1c85-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f1c85-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="f1c85-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="f1c85-141">Composition</span><span class="sxs-lookup"><span data-stu-id="f1c85-141">Compose</span></span><br><span data-ttu-id="f1c85-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-142">Read</span></span> | <span data-ttu-id="f1c85-143">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-143">String</span></span> | [<span data-ttu-id="f1c85-144">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a><span data-ttu-id="f1c85-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="f1c85-145">Namespaces</span></span>

<span data-ttu-id="f1c85-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.</span><span class="sxs-lookup"><span data-stu-id="f1c85-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.4): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="f1c85-147">Détails de l’énumération</span><span class="sxs-lookup"><span data-stu-id="f1c85-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="f1c85-148">AsyncResultStatus : chaîne</span><span class="sxs-lookup"><span data-stu-id="f1c85-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="f1c85-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f1c85-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f1c85-150">Type</span><span class="sxs-lookup"><span data-stu-id="f1c85-150">Type</span></span>

*   <span data-ttu-id="f1c85-151">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f1c85-152">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f1c85-152">Properties:</span></span>

|<span data-ttu-id="f1c85-153">Nom</span><span class="sxs-lookup"><span data-stu-id="f1c85-153">Name</span></span>| <span data-ttu-id="f1c85-154">Type</span><span class="sxs-lookup"><span data-stu-id="f1c85-154">Type</span></span>| <span data-ttu-id="f1c85-155">Description</span><span class="sxs-lookup"><span data-stu-id="f1c85-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f1c85-156">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-156">String</span></span>|<span data-ttu-id="f1c85-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="f1c85-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f1c85-158">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-158">String</span></span>|<span data-ttu-id="f1c85-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="f1c85-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1c85-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f1c85-160">Requirements</span></span>

|<span data-ttu-id="f1c85-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f1c85-161">Requirement</span></span>| <span data-ttu-id="f1c85-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="f1c85-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1c85-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f1c85-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f1c85-164">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-164">1.1</span></span>|
|[<span data-ttu-id="f1c85-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f1c85-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f1c85-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="f1c85-167">CoercionType : chaîne</span><span class="sxs-lookup"><span data-stu-id="f1c85-167">CoercionType: String</span></span>

<span data-ttu-id="f1c85-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f1c85-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f1c85-169">Type</span><span class="sxs-lookup"><span data-stu-id="f1c85-169">Type</span></span>

*   <span data-ttu-id="f1c85-170">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f1c85-171">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f1c85-171">Properties:</span></span>

|<span data-ttu-id="f1c85-172">Nom</span><span class="sxs-lookup"><span data-stu-id="f1c85-172">Name</span></span>| <span data-ttu-id="f1c85-173">Type</span><span class="sxs-lookup"><span data-stu-id="f1c85-173">Type</span></span>| <span data-ttu-id="f1c85-174">Description</span><span class="sxs-lookup"><span data-stu-id="f1c85-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f1c85-175">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-175">String</span></span>|<span data-ttu-id="f1c85-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="f1c85-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f1c85-177">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-177">String</span></span>|<span data-ttu-id="f1c85-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="f1c85-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1c85-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f1c85-179">Requirements</span></span>

|<span data-ttu-id="f1c85-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f1c85-180">Requirement</span></span>| <span data-ttu-id="f1c85-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="f1c85-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1c85-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f1c85-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f1c85-183">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-183">1.1</span></span>|
|[<span data-ttu-id="f1c85-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f1c85-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f1c85-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="f1c85-186">SourceProperty : chaîne</span><span class="sxs-lookup"><span data-stu-id="f1c85-186">SourceProperty: String</span></span>

<span data-ttu-id="f1c85-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="f1c85-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f1c85-188">Type</span><span class="sxs-lookup"><span data-stu-id="f1c85-188">Type</span></span>

*   <span data-ttu-id="f1c85-189">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f1c85-190">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="f1c85-190">Properties:</span></span>

|<span data-ttu-id="f1c85-191">Nom</span><span class="sxs-lookup"><span data-stu-id="f1c85-191">Name</span></span>| <span data-ttu-id="f1c85-192">Type</span><span class="sxs-lookup"><span data-stu-id="f1c85-192">Type</span></span>| <span data-ttu-id="f1c85-193">Description</span><span class="sxs-lookup"><span data-stu-id="f1c85-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f1c85-194">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-194">String</span></span>|<span data-ttu-id="f1c85-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="f1c85-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f1c85-196">String</span><span class="sxs-lookup"><span data-stu-id="f1c85-196">String</span></span>|<span data-ttu-id="f1c85-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="f1c85-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f1c85-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="f1c85-198">Requirements</span></span>

|<span data-ttu-id="f1c85-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="f1c85-199">Requirement</span></span>| <span data-ttu-id="f1c85-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="f1c85-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="f1c85-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="f1c85-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f1c85-202">1.1</span><span class="sxs-lookup"><span data-stu-id="f1c85-202">1.1</span></span>|
|[<span data-ttu-id="f1c85-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="f1c85-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f1c85-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="f1c85-204">Compose or Read</span></span>|
