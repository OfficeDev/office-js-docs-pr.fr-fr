---
title: Office.context - ensemble de conditions requises 1.7
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.7.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: b3dc2442ab418682ac46ad0e1992d561eca98f33
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590819"
---
# <a name="context-mailbox-requirement-set-17"></a><span data-ttu-id="7ac39-103">contexte (ensemble de conditions requises de boîte aux lettres 1.7)</span><span class="sxs-lookup"><span data-stu-id="7ac39-103">context (Mailbox requirement set 1.7)</span></span>

### <a name="officecontext"></a><span data-ttu-id="7ac39-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="7ac39-104">[Office](office.md).context</span></span>

<span data-ttu-id="7ac39-105">Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications.</span><span class="sxs-lookup"><span data-stu-id="7ac39-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="7ac39-106">Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir la liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="7ac39-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.7&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac39-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-107">Requirements</span></span>

|<span data-ttu-id="7ac39-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-108">Requirement</span></span>| <span data-ttu-id="7ac39-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-111">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-111">1.1</span></span>|
|[<span data-ttu-id="7ac39-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="7ac39-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="7ac39-114">Properties</span></span>

| <span data-ttu-id="7ac39-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="7ac39-115">Property</span></span> | <span data-ttu-id="7ac39-116">Modes</span><span class="sxs-lookup"><span data-stu-id="7ac39-116">Modes</span></span> | <span data-ttu-id="7ac39-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="7ac39-117">Return type</span></span> | <span data-ttu-id="7ac39-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="7ac39-118">Minimum</span></span><br><span data-ttu-id="7ac39-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7ac39-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="7ac39-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="7ac39-121">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-121">Compose</span></span><br><span data-ttu-id="7ac39-122">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-122">Read</span></span> | <span data-ttu-id="7ac39-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7ac39-123">String</span></span> | [<span data-ttu-id="7ac39-124">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7ac39-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="7ac39-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="7ac39-126">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-126">Compose</span></span><br><span data-ttu-id="7ac39-127">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-127">Read</span></span> | [<span data-ttu-id="7ac39-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7ac39-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="7ac39-129">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7ac39-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="7ac39-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="7ac39-131">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-131">Compose</span></span><br><span data-ttu-id="7ac39-132">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-132">Read</span></span> | <span data-ttu-id="7ac39-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7ac39-133">String</span></span> | [<span data-ttu-id="7ac39-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7ac39-135">host</span><span class="sxs-lookup"><span data-stu-id="7ac39-135">host</span></span>](#host-hosttype) | <span data-ttu-id="7ac39-136">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-136">Compose</span></span><br><span data-ttu-id="7ac39-137">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-137">Read</span></span> | [<span data-ttu-id="7ac39-138">HostType</span><span class="sxs-lookup"><span data-stu-id="7ac39-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="7ac39-139">1.5</span><span class="sxs-lookup"><span data-stu-id="7ac39-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7ac39-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="7ac39-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="7ac39-141">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-141">Compose</span></span><br><span data-ttu-id="7ac39-142">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-142">Read</span></span> | [<span data-ttu-id="7ac39-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="7ac39-144">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7ac39-145">platform</span><span class="sxs-lookup"><span data-stu-id="7ac39-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="7ac39-146">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-146">Compose</span></span><br><span data-ttu-id="7ac39-147">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-147">Read</span></span> | [<span data-ttu-id="7ac39-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7ac39-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="7ac39-149">1.5</span><span class="sxs-lookup"><span data-stu-id="7ac39-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7ac39-150">requirements</span><span class="sxs-lookup"><span data-stu-id="7ac39-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="7ac39-151">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-151">Compose</span></span><br><span data-ttu-id="7ac39-152">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-152">Read</span></span> | [<span data-ttu-id="7ac39-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7ac39-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="7ac39-154">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7ac39-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="7ac39-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="7ac39-156">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-156">Compose</span></span><br><span data-ttu-id="7ac39-157">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-157">Read</span></span> | [<span data-ttu-id="7ac39-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7ac39-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="7ac39-159">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7ac39-160">ui</span><span class="sxs-lookup"><span data-stu-id="7ac39-160">ui</span></span>](#ui-ui) | <span data-ttu-id="7ac39-161">Composition</span><span class="sxs-lookup"><span data-stu-id="7ac39-161">Compose</span></span><br><span data-ttu-id="7ac39-162">Lire</span><span class="sxs-lookup"><span data-stu-id="7ac39-162">Read</span></span> | [<span data-ttu-id="7ac39-163">UI</span><span class="sxs-lookup"><span data-stu-id="7ac39-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.7&preserve-view=true) | [<span data-ttu-id="7ac39-164">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="7ac39-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="7ac39-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="7ac39-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="7ac39-166">contentLanguage: String</span></span>

<span data-ttu-id="7ac39-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7ac39-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="7ac39-168">La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options de > langue dans l Office `contentLanguage` application cliente.  </span><span class="sxs-lookup"><span data-stu-id="7ac39-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-169">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-169">Type</span></span>

*   <span data-ttu-id="7ac39-170">String</span><span class="sxs-lookup"><span data-stu-id="7ac39-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac39-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-171">Requirements</span></span>

|<span data-ttu-id="7ac39-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-172">Requirement</span></span>| <span data-ttu-id="7ac39-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-175">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-175">1.1</span></span>|
|[<span data-ttu-id="7ac39-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac39-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac39-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="7ac39-179">diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="7ac39-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="7ac39-180">Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="7ac39-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-181">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-181">Type</span></span>

*   [<span data-ttu-id="7ac39-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7ac39-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="7ac39-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-183">Requirements</span></span>

|<span data-ttu-id="7ac39-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-184">Requirement</span></span>| <span data-ttu-id="7ac39-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-187">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-187">1.1</span></span>|
|[<span data-ttu-id="7ac39-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac39-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac39-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="7ac39-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="7ac39-191">displayLanguage: String</span></span>

<span data-ttu-id="7ac39-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.</span><span class="sxs-lookup"><span data-stu-id="7ac39-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="7ac39-193">La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office..  </span><span class="sxs-lookup"><span data-stu-id="7ac39-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-194">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-194">Type</span></span>

*   <span data-ttu-id="7ac39-195">String</span><span class="sxs-lookup"><span data-stu-id="7ac39-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7ac39-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-196">Requirements</span></span>

|<span data-ttu-id="7ac39-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-197">Requirement</span></span>| <span data-ttu-id="7ac39-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-200">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-200">1.1</span></span>|
|[<span data-ttu-id="7ac39-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac39-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac39-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="7ac39-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="7ac39-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="7ac39-205">Obtient Office application qui héberge le module.</span><span class="sxs-lookup"><span data-stu-id="7ac39-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="7ac39-206">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="7ac39-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-207">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-207">Type</span></span>

*   [<span data-ttu-id="7ac39-208">HostType</span><span class="sxs-lookup"><span data-stu-id="7ac39-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="7ac39-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-209">Requirements</span></span>

|<span data-ttu-id="7ac39-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-210">Requirement</span></span>| <span data-ttu-id="7ac39-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-213">1,5</span><span class="sxs-lookup"><span data-stu-id="7ac39-213">1.5</span></span>|
|[<span data-ttu-id="7ac39-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-215">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac39-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac39-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="7ac39-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="7ac39-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="7ac39-218">Fournit la plateforme sur laquelle le module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="7ac39-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="7ac39-219">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="7ac39-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-220">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-220">Type</span></span>

*   [<span data-ttu-id="7ac39-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7ac39-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="7ac39-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-222">Requirements</span></span>

|<span data-ttu-id="7ac39-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-223">Requirement</span></span>| <span data-ttu-id="7ac39-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-226">1,5</span><span class="sxs-lookup"><span data-stu-id="7ac39-226">1.5</span></span>|
|[<span data-ttu-id="7ac39-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac39-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac39-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="7ac39-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="7ac39-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="7ac39-231">Fournit une méthode pour déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="7ac39-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-232">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-232">Type</span></span>

*   [<span data-ttu-id="7ac39-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7ac39-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="7ac39-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-234">Requirements</span></span>

|<span data-ttu-id="7ac39-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-235">Requirement</span></span>| <span data-ttu-id="7ac39-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-238">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-238">1.1</span></span>|
|[<span data-ttu-id="7ac39-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-240">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7ac39-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="7ac39-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="7ac39-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="7ac39-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="7ac39-243">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7ac39-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="7ac39-244">L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.</span><span class="sxs-lookup"><span data-stu-id="7ac39-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-245">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-245">Type</span></span>

*   [<span data-ttu-id="7ac39-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7ac39-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="7ac39-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-247">Requirements</span></span>

|<span data-ttu-id="7ac39-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-248">Requirement</span></span>| <span data-ttu-id="7ac39-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-251">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-251">1.1</span></span>|
|[<span data-ttu-id="7ac39-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7ac39-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="7ac39-253">Restreinte</span><span class="sxs-lookup"><span data-stu-id="7ac39-253">Restricted</span></span>|
|[<span data-ttu-id="7ac39-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="7ac39-256">Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="7ac39-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="7ac39-257">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.</span><span class="sxs-lookup"><span data-stu-id="7ac39-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="7ac39-258">Type</span><span class="sxs-lookup"><span data-stu-id="7ac39-258">Type</span></span>

*   [<span data-ttu-id="7ac39-259">UI</span><span class="sxs-lookup"><span data-stu-id="7ac39-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="7ac39-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7ac39-260">Requirements</span></span>

|<span data-ttu-id="7ac39-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7ac39-261">Requirement</span></span>| <span data-ttu-id="7ac39-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="7ac39-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="7ac39-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7ac39-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7ac39-264">1.1</span><span class="sxs-lookup"><span data-stu-id="7ac39-264">1.1</span></span>|
|[<span data-ttu-id="7ac39-265">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7ac39-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7ac39-266">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7ac39-266">Compose or Read</span></span>|
