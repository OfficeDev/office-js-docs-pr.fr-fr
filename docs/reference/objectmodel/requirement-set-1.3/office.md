---
title: Office de noms - ensemble de conditions requises 1.3
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.3.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: f4aecf016e259141fd8adb2683864d4c36bdaf4b
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592002"
---
# <a name="office-mailbox-requirement-set-13"></a><span data-ttu-id="d8312-103">Office (ensemble de conditions requises de boîte aux lettres 1.3)</span><span class="sxs-lookup"><span data-stu-id="d8312-103">Office (Mailbox requirement set 1.3)</span></span>

<span data-ttu-id="d8312-p101">L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d8312-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Common API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d8312-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d8312-106">Requirements</span></span>

|<span data-ttu-id="d8312-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d8312-107">Requirement</span></span>| <span data-ttu-id="d8312-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="d8312-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8312-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d8312-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d8312-110">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-110">1.1</span></span>|
|[<span data-ttu-id="d8312-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d8312-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d8312-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d8312-112">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="d8312-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="d8312-113">Properties</span></span>

| <span data-ttu-id="d8312-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="d8312-114">Property</span></span> | <span data-ttu-id="d8312-115">Modes</span><span class="sxs-lookup"><span data-stu-id="d8312-115">Modes</span></span> | <span data-ttu-id="d8312-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="d8312-116">Return type</span></span> | <span data-ttu-id="d8312-117">Minimum</span><span class="sxs-lookup"><span data-stu-id="d8312-117">Minimum</span></span><br><span data-ttu-id="d8312-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="d8312-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d8312-119">context</span><span class="sxs-lookup"><span data-stu-id="d8312-119">context</span></span>](office.context.md) | <span data-ttu-id="d8312-120">Composition</span><span class="sxs-lookup"><span data-stu-id="d8312-120">Compose</span></span><br><span data-ttu-id="d8312-121">Lire</span><span class="sxs-lookup"><span data-stu-id="d8312-121">Read</span></span> | [<span data-ttu-id="d8312-122">Context</span><span class="sxs-lookup"><span data-stu-id="d8312-122">Context</span></span>](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="d8312-123">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a><span data-ttu-id="d8312-124">Énumérations</span><span class="sxs-lookup"><span data-stu-id="d8312-124">Enumerations</span></span>

| <span data-ttu-id="d8312-125">Énumération</span><span class="sxs-lookup"><span data-stu-id="d8312-125">Enumeration</span></span> | <span data-ttu-id="d8312-126">Modes</span><span class="sxs-lookup"><span data-stu-id="d8312-126">Modes</span></span> | <span data-ttu-id="d8312-127">Type de retour</span><span class="sxs-lookup"><span data-stu-id="d8312-127">Return type</span></span> | <span data-ttu-id="d8312-128">Minimum</span><span class="sxs-lookup"><span data-stu-id="d8312-128">Minimum</span></span><br><span data-ttu-id="d8312-129">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="d8312-129">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d8312-130">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d8312-130">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d8312-131">Composition</span><span class="sxs-lookup"><span data-stu-id="d8312-131">Compose</span></span><br><span data-ttu-id="d8312-132">Lire</span><span class="sxs-lookup"><span data-stu-id="d8312-132">Read</span></span> | <span data-ttu-id="d8312-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d8312-133">String</span></span> | [<span data-ttu-id="d8312-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d8312-135">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d8312-135">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d8312-136">Composition</span><span class="sxs-lookup"><span data-stu-id="d8312-136">Compose</span></span><br><span data-ttu-id="d8312-137">Lire</span><span class="sxs-lookup"><span data-stu-id="d8312-137">Read</span></span> | <span data-ttu-id="d8312-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d8312-138">String</span></span> | [<span data-ttu-id="d8312-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d8312-140">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d8312-140">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d8312-141">Composition</span><span class="sxs-lookup"><span data-stu-id="d8312-141">Compose</span></span><br><span data-ttu-id="d8312-142">Lire</span><span class="sxs-lookup"><span data-stu-id="d8312-142">Read</span></span> | <span data-ttu-id="d8312-143">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d8312-143">String</span></span> | [<span data-ttu-id="d8312-144">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a><span data-ttu-id="d8312-145">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="d8312-145">Namespaces</span></span>

<span data-ttu-id="d8312-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .</span><span class="sxs-lookup"><span data-stu-id="d8312-146">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.3&preserve-view=true): Includes a number of Outlook-specific enumerations, for example, `ItemType`, `EntityType`, `AttachmentType`, `RecipientType`, `ResponseType`, and `ItemNotificationMessageType`.</span></span>

## <a name="enumeration-details"></a><span data-ttu-id="d8312-147">Détails de l’éumération</span><span class="sxs-lookup"><span data-stu-id="d8312-147">Enumeration details</span></span>

#### <a name="asyncresultstatus-string"></a><span data-ttu-id="d8312-148">AsyncResultStatus: String</span><span class="sxs-lookup"><span data-stu-id="d8312-148">AsyncResultStatus: String</span></span>

<span data-ttu-id="d8312-149">Spécifie le résultat d’un appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="d8312-149">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d8312-150">Type</span><span class="sxs-lookup"><span data-stu-id="d8312-150">Type</span></span>

*   <span data-ttu-id="d8312-151">String</span><span class="sxs-lookup"><span data-stu-id="d8312-151">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d8312-152">Propriétés</span><span class="sxs-lookup"><span data-stu-id="d8312-152">Properties</span></span>

|<span data-ttu-id="d8312-153">Nom</span><span class="sxs-lookup"><span data-stu-id="d8312-153">Name</span></span>| <span data-ttu-id="d8312-154">Type</span><span class="sxs-lookup"><span data-stu-id="d8312-154">Type</span></span>| <span data-ttu-id="d8312-155">Description</span><span class="sxs-lookup"><span data-stu-id="d8312-155">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d8312-156">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d8312-156">String</span></span>|<span data-ttu-id="d8312-157">L’appel a réussi.</span><span class="sxs-lookup"><span data-stu-id="d8312-157">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d8312-158">String</span><span class="sxs-lookup"><span data-stu-id="d8312-158">String</span></span>|<span data-ttu-id="d8312-159">L’appel n’a pas réussi.</span><span class="sxs-lookup"><span data-stu-id="d8312-159">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8312-160">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d8312-160">Requirements</span></span>

|<span data-ttu-id="d8312-161">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d8312-161">Requirement</span></span>| <span data-ttu-id="d8312-162">Valeur</span><span class="sxs-lookup"><span data-stu-id="d8312-162">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8312-163">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d8312-163">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d8312-164">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-164">1.1</span></span>|
|[<span data-ttu-id="d8312-165">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d8312-165">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d8312-166">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d8312-166">Compose or Read</span></span>|

<br>

---
---

#### <a name="coerciontype-string"></a><span data-ttu-id="d8312-167">CoercionType: String</span><span class="sxs-lookup"><span data-stu-id="d8312-167">CoercionType: String</span></span>

<span data-ttu-id="d8312-168">Indique comment forcer le type des données retournées ou définies par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d8312-168">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d8312-169">Type</span><span class="sxs-lookup"><span data-stu-id="d8312-169">Type</span></span>

*   <span data-ttu-id="d8312-170">String</span><span class="sxs-lookup"><span data-stu-id="d8312-170">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d8312-171">Propriétés</span><span class="sxs-lookup"><span data-stu-id="d8312-171">Properties</span></span>

|<span data-ttu-id="d8312-172">Nom</span><span class="sxs-lookup"><span data-stu-id="d8312-172">Name</span></span>| <span data-ttu-id="d8312-173">Type</span><span class="sxs-lookup"><span data-stu-id="d8312-173">Type</span></span>| <span data-ttu-id="d8312-174">Description</span><span class="sxs-lookup"><span data-stu-id="d8312-174">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d8312-175">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d8312-175">String</span></span>|<span data-ttu-id="d8312-176">Demande que les données soient renvoyées au format HTML.</span><span class="sxs-lookup"><span data-stu-id="d8312-176">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d8312-177">String</span><span class="sxs-lookup"><span data-stu-id="d8312-177">String</span></span>|<span data-ttu-id="d8312-178">Demande que les données soient renvoyées au format texte.</span><span class="sxs-lookup"><span data-stu-id="d8312-178">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8312-179">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d8312-179">Requirements</span></span>

|<span data-ttu-id="d8312-180">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d8312-180">Requirement</span></span>| <span data-ttu-id="d8312-181">Valeur</span><span class="sxs-lookup"><span data-stu-id="d8312-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8312-182">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d8312-182">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d8312-183">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-183">1.1</span></span>|
|[<span data-ttu-id="d8312-184">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d8312-184">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d8312-185">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d8312-185">Compose or Read</span></span>|

<br>

---
---

#### <a name="sourceproperty-string"></a><span data-ttu-id="d8312-186">SourceProperty: String</span><span class="sxs-lookup"><span data-stu-id="d8312-186">SourceProperty: String</span></span>

<span data-ttu-id="d8312-187">Spécifie la source des données renvoyées par la méthode appelée.</span><span class="sxs-lookup"><span data-stu-id="d8312-187">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d8312-188">Type</span><span class="sxs-lookup"><span data-stu-id="d8312-188">Type</span></span>

*   <span data-ttu-id="d8312-189">String</span><span class="sxs-lookup"><span data-stu-id="d8312-189">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d8312-190">Propriétés</span><span class="sxs-lookup"><span data-stu-id="d8312-190">Properties</span></span>

|<span data-ttu-id="d8312-191">Nom</span><span class="sxs-lookup"><span data-stu-id="d8312-191">Name</span></span>| <span data-ttu-id="d8312-192">Type</span><span class="sxs-lookup"><span data-stu-id="d8312-192">Type</span></span>| <span data-ttu-id="d8312-193">Description</span><span class="sxs-lookup"><span data-stu-id="d8312-193">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d8312-194">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d8312-194">String</span></span>|<span data-ttu-id="d8312-195">La source de données est dans le corps d’un message.</span><span class="sxs-lookup"><span data-stu-id="d8312-195">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d8312-196">String</span><span class="sxs-lookup"><span data-stu-id="d8312-196">String</span></span>|<span data-ttu-id="d8312-197">La source de données est dans l’objet d’un message.</span><span class="sxs-lookup"><span data-stu-id="d8312-197">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d8312-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d8312-198">Requirements</span></span>

|<span data-ttu-id="d8312-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d8312-199">Requirement</span></span>| <span data-ttu-id="d8312-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="d8312-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="d8312-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d8312-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d8312-202">1.1</span><span class="sxs-lookup"><span data-stu-id="d8312-202">1.1</span></span>|
|[<span data-ttu-id="d8312-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d8312-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d8312-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d8312-204">Compose or Read</span></span>|
