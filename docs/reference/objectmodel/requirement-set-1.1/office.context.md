---
title: Office. Context-ensemble de conditions requises 1,1
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,1.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 2474c5f31dcd996785f61f06528ffcf3a92b05c2
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430561"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="6cb59-103">contexte (boîte aux lettres requise définie sur 1,1)</span><span class="sxs-lookup"><span data-stu-id="6cb59-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="6cb59-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="6cb59-104">[Office](office.md).context</span></span>

<span data-ttu-id="6cb59-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="6cb59-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="6cb59-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="6cb59-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cb59-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-107">Requirements</span></span>

|<span data-ttu-id="6cb59-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-108">Requirement</span></span>| <span data-ttu-id="6cb59-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-111">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-111">1.1</span></span>|
|[<span data-ttu-id="6cb59-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6cb59-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="6cb59-114">Properties</span></span>

| <span data-ttu-id="6cb59-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="6cb59-115">Property</span></span> | <span data-ttu-id="6cb59-116">Modes</span><span class="sxs-lookup"><span data-stu-id="6cb59-116">Modes</span></span> | <span data-ttu-id="6cb59-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="6cb59-117">Return type</span></span> | <span data-ttu-id="6cb59-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="6cb59-118">Minimum</span></span><br><span data-ttu-id="6cb59-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6cb59-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="6cb59-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="6cb59-121">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-121">Compose</span></span><br><span data-ttu-id="6cb59-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-122">Read</span></span> | <span data-ttu-id="6cb59-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cb59-123">String</span></span> | [<span data-ttu-id="6cb59-124">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="6cb59-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="6cb59-126">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-126">Compose</span></span><br><span data-ttu-id="6cb59-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-127">Read</span></span> | [<span data-ttu-id="6cb59-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6cb59-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="6cb59-129">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="6cb59-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="6cb59-131">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-131">Compose</span></span><br><span data-ttu-id="6cb59-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-132">Read</span></span> | <span data-ttu-id="6cb59-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6cb59-133">String</span></span> | [<span data-ttu-id="6cb59-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-135">hote</span><span class="sxs-lookup"><span data-stu-id="6cb59-135">host</span></span>](#host-hosttype) | <span data-ttu-id="6cb59-136">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-136">Compose</span></span><br><span data-ttu-id="6cb59-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-137">Read</span></span> | [<span data-ttu-id="6cb59-138">HostType</span><span class="sxs-lookup"><span data-stu-id="6cb59-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="6cb59-139">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="6cb59-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="6cb59-141">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-141">Compose</span></span><br><span data-ttu-id="6cb59-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-142">Read</span></span> | [<span data-ttu-id="6cb59-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="6cb59-144">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-145">plateforme</span><span class="sxs-lookup"><span data-stu-id="6cb59-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="6cb59-146">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-146">Compose</span></span><br><span data-ttu-id="6cb59-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-147">Read</span></span> | [<span data-ttu-id="6cb59-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6cb59-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="6cb59-149">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-150">requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="6cb59-151">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-151">Compose</span></span><br><span data-ttu-id="6cb59-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-152">Read</span></span> | [<span data-ttu-id="6cb59-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6cb59-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="6cb59-154">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="6cb59-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="6cb59-156">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-156">Compose</span></span><br><span data-ttu-id="6cb59-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-157">Read</span></span> | [<span data-ttu-id="6cb59-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6cb59-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="6cb59-159">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6cb59-160">ui</span><span class="sxs-lookup"><span data-stu-id="6cb59-160">ui</span></span>](#ui-ui) | <span data-ttu-id="6cb59-161">Composition</span><span class="sxs-lookup"><span data-stu-id="6cb59-161">Compose</span></span><br><span data-ttu-id="6cb59-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-162">Read</span></span> | [<span data-ttu-id="6cb59-163">UI</span><span class="sxs-lookup"><span data-stu-id="6cb59-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="6cb59-164">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="6cb59-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="6cb59-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="6cb59-166">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="6cb59-166">contentLanguage: String</span></span>

<span data-ttu-id="6cb59-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6cb59-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="6cb59-168">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="6cb59-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-169">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-169">Type</span></span>

*   <span data-ttu-id="6cb59-170">String</span><span class="sxs-lookup"><span data-stu-id="6cb59-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cb59-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-171">Requirements</span></span>

|<span data-ttu-id="6cb59-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-172">Requirement</span></span>| <span data-ttu-id="6cb59-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-175">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-175">1.1</span></span>|
|[<span data-ttu-id="6cb59-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb59-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cb59-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="6cb59-179">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6cb59-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="6cb59-180">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="6cb59-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-181">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-181">Type</span></span>

*   [<span data-ttu-id="6cb59-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6cb59-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="6cb59-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-183">Requirements</span></span>

|<span data-ttu-id="6cb59-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-184">Requirement</span></span>| <span data-ttu-id="6cb59-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-187">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-187">1.1</span></span>|
|[<span data-ttu-id="6cb59-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb59-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cb59-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="6cb59-191">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="6cb59-191">displayLanguage: String</span></span>

<span data-ttu-id="6cb59-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="6cb59-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="6cb59-193">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="6cb59-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-194">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-194">Type</span></span>

*   <span data-ttu-id="6cb59-195">String</span><span class="sxs-lookup"><span data-stu-id="6cb59-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6cb59-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-196">Requirements</span></span>

|<span data-ttu-id="6cb59-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-197">Requirement</span></span>| <span data-ttu-id="6cb59-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-200">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-200">1.1</span></span>|
|[<span data-ttu-id="6cb59-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb59-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cb59-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="6cb59-204">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="6cb59-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="6cb59-205">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="6cb59-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-206">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-206">Type</span></span>

*   [<span data-ttu-id="6cb59-207">HostType</span><span class="sxs-lookup"><span data-stu-id="6cb59-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="6cb59-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-208">Requirements</span></span>

|<span data-ttu-id="6cb59-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-209">Requirement</span></span>| <span data-ttu-id="6cb59-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-212">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-212">1.1</span></span>|
|[<span data-ttu-id="6cb59-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-214">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb59-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cb59-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="6cb59-216">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="6cb59-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="6cb59-217">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="6cb59-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-218">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-218">Type</span></span>

*   [<span data-ttu-id="6cb59-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6cb59-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="6cb59-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-220">Requirements</span></span>

|<span data-ttu-id="6cb59-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-221">Requirement</span></span>| <span data-ttu-id="6cb59-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-224">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-224">1.1</span></span>|
|[<span data-ttu-id="6cb59-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb59-227">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cb59-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="6cb59-228">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="6cb59-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="6cb59-229">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="6cb59-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-230">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-230">Type</span></span>

*   [<span data-ttu-id="6cb59-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6cb59-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="6cb59-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-232">Requirements</span></span>

|<span data-ttu-id="6cb59-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-233">Requirement</span></span>| <span data-ttu-id="6cb59-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-236">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-236">1.1</span></span>|
|[<span data-ttu-id="6cb59-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6cb59-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="6cb59-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="6cb59-240">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="6cb59-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="6cb59-241">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6cb59-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6cb59-242">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="6cb59-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-243">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-243">Type</span></span>

*   [<span data-ttu-id="6cb59-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6cb59-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6cb59-245">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-245">Requirements</span></span>

|<span data-ttu-id="6cb59-246">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-246">Requirement</span></span>| <span data-ttu-id="6cb59-247">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-248">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-249">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-249">1.1</span></span>|
|[<span data-ttu-id="6cb59-250">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6cb59-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="6cb59-251">Restreinte</span><span class="sxs-lookup"><span data-stu-id="6cb59-251">Restricted</span></span>|
|[<span data-ttu-id="6cb59-252">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-253">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="6cb59-254">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="6cb59-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="6cb59-255">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="6cb59-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6cb59-256">Type</span><span class="sxs-lookup"><span data-stu-id="6cb59-256">Type</span></span>

*   [<span data-ttu-id="6cb59-257">UI</span><span class="sxs-lookup"><span data-stu-id="6cb59-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="6cb59-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6cb59-258">Requirements</span></span>

|<span data-ttu-id="6cb59-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6cb59-259">Requirement</span></span>| <span data-ttu-id="6cb59-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="6cb59-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="6cb59-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6cb59-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6cb59-262">1.1</span><span class="sxs-lookup"><span data-stu-id="6cb59-262">1.1</span></span>|
|[<span data-ttu-id="6cb59-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6cb59-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6cb59-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6cb59-264">Compose or Read</span></span>|
