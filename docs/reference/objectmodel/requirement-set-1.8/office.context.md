---
title: Office.context - ensemble de conditions requises 1.8
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.8.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 99573d9984c571c99461e90e8bdccdca35fe30b7
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590966"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="ac87c-103">context (ensemble de conditions requises de boîte aux lettres 1.8)</span><span class="sxs-lookup"><span data-stu-id="ac87c-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ac87c-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ac87c-104">[Office](office.md).context</span></span>

<span data-ttu-id="ac87c-105">Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications.</span><span class="sxs-lookup"><span data-stu-id="ac87c-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ac87c-106">Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir la liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="ac87c-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac87c-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-107">Requirements</span></span>

|<span data-ttu-id="ac87c-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-108">Requirement</span></span>| <span data-ttu-id="ac87c-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-111">1.1</span></span>|
|[<span data-ttu-id="ac87c-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="ac87c-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ac87c-114">Properties</span></span>

| <span data-ttu-id="ac87c-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="ac87c-115">Property</span></span> | <span data-ttu-id="ac87c-116">Modes</span><span class="sxs-lookup"><span data-stu-id="ac87c-116">Modes</span></span> | <span data-ttu-id="ac87c-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ac87c-117">Return type</span></span> | <span data-ttu-id="ac87c-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="ac87c-118">Minimum</span></span><br><span data-ttu-id="ac87c-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ac87c-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ac87c-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ac87c-121">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-121">Compose</span></span><br><span data-ttu-id="ac87c-122">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-122">Read</span></span> | <span data-ttu-id="ac87c-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ac87c-123">String</span></span> | [<span data-ttu-id="ac87c-124">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac87c-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="ac87c-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ac87c-126">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-126">Compose</span></span><br><span data-ttu-id="ac87c-127">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-127">Read</span></span> | [<span data-ttu-id="ac87c-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ac87c-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="ac87c-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac87c-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ac87c-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ac87c-131">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-131">Compose</span></span><br><span data-ttu-id="ac87c-132">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-132">Read</span></span> | <span data-ttu-id="ac87c-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ac87c-133">String</span></span> | [<span data-ttu-id="ac87c-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac87c-135">host</span><span class="sxs-lookup"><span data-stu-id="ac87c-135">host</span></span>](#host-hosttype) | <span data-ttu-id="ac87c-136">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-136">Compose</span></span><br><span data-ttu-id="ac87c-137">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-137">Read</span></span> | [<span data-ttu-id="ac87c-138">HostType</span><span class="sxs-lookup"><span data-stu-id="ac87c-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="ac87c-139">1.5</span><span class="sxs-lookup"><span data-stu-id="ac87c-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ac87c-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="ac87c-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ac87c-141">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-141">Compose</span></span><br><span data-ttu-id="ac87c-142">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-142">Read</span></span> | [<span data-ttu-id="ac87c-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="ac87c-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac87c-145">platform</span><span class="sxs-lookup"><span data-stu-id="ac87c-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ac87c-146">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-146">Compose</span></span><br><span data-ttu-id="ac87c-147">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-147">Read</span></span> | [<span data-ttu-id="ac87c-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ac87c-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="ac87c-149">1.5</span><span class="sxs-lookup"><span data-stu-id="ac87c-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="ac87c-150">requirements</span><span class="sxs-lookup"><span data-stu-id="ac87c-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ac87c-151">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-151">Compose</span></span><br><span data-ttu-id="ac87c-152">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-152">Read</span></span> | [<span data-ttu-id="ac87c-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ac87c-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="ac87c-154">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac87c-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ac87c-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ac87c-156">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-156">Compose</span></span><br><span data-ttu-id="ac87c-157">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-157">Read</span></span> | [<span data-ttu-id="ac87c-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ac87c-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="ac87c-159">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ac87c-160">ui</span><span class="sxs-lookup"><span data-stu-id="ac87c-160">ui</span></span>](#ui-ui) | <span data-ttu-id="ac87c-161">Composition</span><span class="sxs-lookup"><span data-stu-id="ac87c-161">Compose</span></span><br><span data-ttu-id="ac87c-162">Lire</span><span class="sxs-lookup"><span data-stu-id="ac87c-162">Read</span></span> | [<span data-ttu-id="ac87c-163">UI</span><span class="sxs-lookup"><span data-stu-id="ac87c-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="ac87c-164">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ac87c-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="ac87c-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="ac87c-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="ac87c-166">contentLanguage: String</span></span>

<span data-ttu-id="ac87c-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="ac87c-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ac87c-168">La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options d'> langue dans `contentLanguage` l’application cliente Office’édition.  </span><span class="sxs-lookup"><span data-stu-id="ac87c-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-169">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-169">Type</span></span>

*   <span data-ttu-id="ac87c-170">String</span><span class="sxs-lookup"><span data-stu-id="ac87c-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac87c-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-171">Requirements</span></span>

|<span data-ttu-id="ac87c-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-172">Requirement</span></span>| <span data-ttu-id="ac87c-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-175">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-175">1.1</span></span>|
|[<span data-ttu-id="ac87c-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac87c-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac87c-178">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ac87c-179">diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ac87c-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ac87c-180">Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ac87c-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-181">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-181">Type</span></span>

*   [<span data-ttu-id="ac87c-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ac87c-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ac87c-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-183">Requirements</span></span>

|<span data-ttu-id="ac87c-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-184">Requirement</span></span>| <span data-ttu-id="ac87c-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-187">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-187">1.1</span></span>|
|[<span data-ttu-id="ac87c-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac87c-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac87c-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ac87c-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="ac87c-191">displayLanguage: String</span></span>

<span data-ttu-id="ac87c-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.</span><span class="sxs-lookup"><span data-stu-id="ac87c-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="ac87c-193">La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office..  </span><span class="sxs-lookup"><span data-stu-id="ac87c-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-194">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-194">Type</span></span>

*   <span data-ttu-id="ac87c-195">String</span><span class="sxs-lookup"><span data-stu-id="ac87c-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ac87c-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-196">Requirements</span></span>

|<span data-ttu-id="ac87c-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-197">Requirement</span></span>| <span data-ttu-id="ac87c-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-200">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-200">1.1</span></span>|
|[<span data-ttu-id="ac87c-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac87c-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac87c-203">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a><span data-ttu-id="ac87c-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ac87c-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ac87c-205">Obtient Office application qui héberge le module.</span><span class="sxs-lookup"><span data-stu-id="ac87c-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="ac87c-206">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="ac87c-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-207">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-207">Type</span></span>

*   [<span data-ttu-id="ac87c-208">HostType</span><span class="sxs-lookup"><span data-stu-id="ac87c-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ac87c-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-209">Requirements</span></span>

|<span data-ttu-id="ac87c-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-210">Requirement</span></span>| <span data-ttu-id="ac87c-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-213">1,5</span><span class="sxs-lookup"><span data-stu-id="ac87c-213">1.5</span></span>|
|[<span data-ttu-id="ac87c-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-215">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac87c-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac87c-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="ac87c-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ac87c-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ac87c-218">Fournit la plateforme sur laquelle le module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ac87c-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="ac87c-219">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="ac87c-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-220">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-220">Type</span></span>

*   [<span data-ttu-id="ac87c-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ac87c-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ac87c-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-222">Requirements</span></span>

|<span data-ttu-id="ac87c-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-223">Requirement</span></span>| <span data-ttu-id="ac87c-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-226">1,5</span><span class="sxs-lookup"><span data-stu-id="ac87c-226">1.5</span></span>|
|[<span data-ttu-id="ac87c-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac87c-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac87c-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ac87c-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ac87c-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ac87c-231">Fournit une méthode pour déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="ac87c-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-232">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-232">Type</span></span>

*   [<span data-ttu-id="ac87c-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ac87c-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ac87c-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-234">Requirements</span></span>

|<span data-ttu-id="ac87c-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-235">Requirement</span></span>| <span data-ttu-id="ac87c-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-238">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-238">1.1</span></span>|
|[<span data-ttu-id="ac87c-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-240">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ac87c-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="ac87c-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ac87c-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ac87c-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ac87c-243">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ac87c-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ac87c-244">L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.</span><span class="sxs-lookup"><span data-stu-id="ac87c-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-245">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-245">Type</span></span>

*   [<span data-ttu-id="ac87c-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ac87c-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ac87c-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-247">Requirements</span></span>

|<span data-ttu-id="ac87c-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-248">Requirement</span></span>| <span data-ttu-id="ac87c-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-251">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-251">1.1</span></span>|
|[<span data-ttu-id="ac87c-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ac87c-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ac87c-253">Restreinte</span><span class="sxs-lookup"><span data-stu-id="ac87c-253">Restricted</span></span>|
|[<span data-ttu-id="ac87c-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ac87c-256">Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ac87c-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ac87c-257">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.</span><span class="sxs-lookup"><span data-stu-id="ac87c-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ac87c-258">Type</span><span class="sxs-lookup"><span data-stu-id="ac87c-258">Type</span></span>

*   [<span data-ttu-id="ac87c-259">UI</span><span class="sxs-lookup"><span data-stu-id="ac87c-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ac87c-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ac87c-260">Requirements</span></span>

|<span data-ttu-id="ac87c-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ac87c-261">Requirement</span></span>| <span data-ttu-id="ac87c-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="ac87c-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="ac87c-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ac87c-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ac87c-264">1.1</span><span class="sxs-lookup"><span data-stu-id="ac87c-264">1.1</span></span>|
|[<span data-ttu-id="ac87c-265">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ac87c-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ac87c-266">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ac87c-266">Compose or Read</span></span>|
