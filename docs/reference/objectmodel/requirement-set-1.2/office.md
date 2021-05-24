---
title: Office de noms - ensemble de conditions requises 1.2
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.2.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 4cd15d77d1c5d9b95152f038f3421c5838bfb84f
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590406"
---
# <a name="office-mailbox-requirement-set-12"></a><span data-ttu-id="b5b90-103">Office (ensemble de conditions requises de boîte aux lettres 1.2)</span><span class="sxs-lookup"><span data-stu-id="b5b90-103">Office (Mailbox requirement set 1.2)</span></span>

<span data-ttu-id="b5b90-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="b5b90-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b5b90-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b5b90-106">Requirements</span></span>

|<span data-ttu-id="b5b90-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b5b90-107">Requirement</span></span>| <span data-ttu-id="b5b90-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="b5b90-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5b90-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b5b90-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5b90-110">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-110">1.1</span></span>|
|[<span data-ttu-id="b5b90-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b5b90-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5b90-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b5b90-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="b5b90-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="b5b90-113">Properties</span></span>

| <span data-ttu-id="b5b90-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="b5b90-114">Property</span></span> | <span data-ttu-id="b5b90-115">Modes</span><span class="sxs-lookup"><span data-stu-id="b5b90-115">Modes</span></span> | <span data-ttu-id="b5b90-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="b5b90-116">Return type</span></span> | <span data-ttu-id="b5b90-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="b5b90-117">Minimum</span></span><br><span data-ttu-id="b5b90-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b5b90-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b5b90-119">context</span><span class="sxs-lookup"><span data-stu-id="b5b90-119">context</span></span>](office.context.md) | <span data-ttu-id="b5b90-120">Composition</span><span class="sxs-lookup"><span data-stu-id="b5b90-120">Compose</span></span><br><span data-ttu-id="b5b90-121">Lire</span><span class="sxs-lookup"><span data-stu-id="b5b90-121">Read</span></span> | [<span data-ttu-id="b5b90-122">Context</span><span class="sxs-lookup"><span data-stu-id="b5b90-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="b5b90-123">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="b5b90-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="b5b90-124">Enumerations</span></span>

| <span data-ttu-id="b5b90-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="b5b90-125">Enumeration</span></span> | <span data-ttu-id="b5b90-126">Modes</span><span class="sxs-lookup"><span data-stu-id="b5b90-126">Modes</span></span> | <span data-ttu-id="b5b90-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="b5b90-127">Return type</span></span> | <span data-ttu-id="b5b90-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="b5b90-128">Minimum</span></span><br><span data-ttu-id="b5b90-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b5b90-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b5b90-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="b5b90-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="b5b90-131">Composition</span><span class="sxs-lookup"><span data-stu-id="b5b90-131">Compose</span></span><br><span data-ttu-id="b5b90-132">Lire</span><span class="sxs-lookup"><span data-stu-id="b5b90-132">Read</span></span> | <span data-ttu-id="b5b90-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b5b90-133">String</span></span> | [<span data-ttu-id="b5b90-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b5b90-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="b5b90-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="b5b90-136">Composition</span><span class="sxs-lookup"><span data-stu-id="b5b90-136">Compose</span></span><br><span data-ttu-id="b5b90-137">Lire</span><span class="sxs-lookup"><span data-stu-id="b5b90-137">Read</span></span> | <span data-ttu-id="b5b90-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b5b90-138">String</span></span> | [<span data-ttu-id="b5b90-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b5b90-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="b5b90-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="b5b90-141">Composition</span><span class="sxs-lookup"><span data-stu-id="b5b90-141">Compose</span></span><br><span data-ttu-id="b5b90-142">Lire</span><span class="sxs-lookup"><span data-stu-id="b5b90-142">Read</span></span> | <span data-ttu-id="b5b90-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b5b90-143">String</span></span> | [<span data-ttu-id="b5b90-144">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="b5b90-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="b5b90-145">Namespaces</span></span>

<span data-ttu-id="b5b90-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="b5b90-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="b5b90-147">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="b5b90-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="b5b90-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="b5b90-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="b5b90-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="b5b90-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="b5b90-150">Type</span><span class="sxs-lookup"><span data-stu-id="b5b90-150">Type</span></span>

*   <span data-ttu-id="b5b90-151">String</span><span class="sxs-lookup"><span data-stu-id="b5b90-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b5b90-152">Propriétés</span><span class="sxs-lookup"><span data-stu-id="b5b90-152">Properties</span></span>

|<span data-ttu-id="b5b90-153">Nom</span><span class="sxs-lookup"><span data-stu-id="b5b90-153">Name</span></span>| <span data-ttu-id="b5b90-154">Type</span><span class="sxs-lookup"><span data-stu-id="b5b90-154">Type</span></span>| <span data-ttu-id="b5b90-155">Description</span><span class="sxs-lookup"><span data-stu-id="b5b90-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="b5b90-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b5b90-156">String</span></span>|<span data-ttu-id="b5b90-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="b5b90-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="b5b90-158">String</span><span class="sxs-lookup"><span data-stu-id="b5b90-158">String</span></span>|<span data-ttu-id="b5b90-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="b5b90-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b5b90-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b5b90-160">Requirements</span></span>

|<span data-ttu-id="b5b90-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b5b90-161">Requirement</span></span>| <span data-ttu-id="b5b90-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="b5b90-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5b90-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b5b90-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5b90-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-164">1.1</span></span>|
|[<span data-ttu-id="b5b90-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b5b90-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5b90-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b5b90-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="b5b90-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="b5b90-167">CoercionType: String</span></span>

<span data-ttu-id="b5b90-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="b5b90-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b5b90-169">Type</span><span class="sxs-lookup"><span data-stu-id="b5b90-169">Type</span></span>

*   <span data-ttu-id="b5b90-170">String</span><span class="sxs-lookup"><span data-stu-id="b5b90-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b5b90-171">Propriétés</span><span class="sxs-lookup"><span data-stu-id="b5b90-171">Properties</span></span>

|<span data-ttu-id="b5b90-172">Nom</span><span class="sxs-lookup"><span data-stu-id="b5b90-172">Name</span></span>| <span data-ttu-id="b5b90-173">Type</span><span class="sxs-lookup"><span data-stu-id="b5b90-173">Type</span></span>| <span data-ttu-id="b5b90-174">Description</span><span class="sxs-lookup"><span data-stu-id="b5b90-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="b5b90-175">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b5b90-175">String</span></span>|<span data-ttu-id="b5b90-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="b5b90-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="b5b90-177">String</span><span class="sxs-lookup"><span data-stu-id="b5b90-177">String</span></span>|<span data-ttu-id="b5b90-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="b5b90-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b5b90-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b5b90-179">Requirements</span></span>

|<span data-ttu-id="b5b90-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b5b90-180">Requirement</span></span>| <span data-ttu-id="b5b90-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="b5b90-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5b90-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b5b90-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5b90-183">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-183">1.1</span></span>|
|[<span data-ttu-id="b5b90-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b5b90-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5b90-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b5b90-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="b5b90-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="b5b90-186">SourceProperty: String</span></span>

<span data-ttu-id="b5b90-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="b5b90-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="b5b90-188">Type</span><span class="sxs-lookup"><span data-stu-id="b5b90-188">Type</span></span>

*   <span data-ttu-id="b5b90-189">String</span><span class="sxs-lookup"><span data-stu-id="b5b90-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="b5b90-190">Propriétés</span><span class="sxs-lookup"><span data-stu-id="b5b90-190">Properties</span></span>

|<span data-ttu-id="b5b90-191">Nom</span><span class="sxs-lookup"><span data-stu-id="b5b90-191">Name</span></span>| <span data-ttu-id="b5b90-192">Type</span><span class="sxs-lookup"><span data-stu-id="b5b90-192">Type</span></span>| <span data-ttu-id="b5b90-193">Description</span><span class="sxs-lookup"><span data-stu-id="b5b90-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="b5b90-194">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b5b90-194">String</span></span>|<span data-ttu-id="b5b90-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="b5b90-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="b5b90-196">String</span><span class="sxs-lookup"><span data-stu-id="b5b90-196">String</span></span>|<span data-ttu-id="b5b90-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="b5b90-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b5b90-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b5b90-198">Requirements</span></span>

|<span data-ttu-id="b5b90-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b5b90-199">Requirement</span></span>| <span data-ttu-id="b5b90-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="b5b90-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="b5b90-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b5b90-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b5b90-202">1.1</span><span class="sxs-lookup"><span data-stu-id="b5b90-202">1.1</span></span>|
|[<span data-ttu-id="b5b90-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b5b90-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b5b90-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b5b90-204">Compose or Read</span></span>|
