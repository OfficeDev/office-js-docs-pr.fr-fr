---
title: Office. Context-ensemble de conditions requises 1,3
description: Modèle objet de l’objet de contexte Outlook dans l’API des compléments Outlook (version 1,3 de l’API de boîte aux lettres).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: dc524c5657380845744e241a600fbf66f1eed5ce
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720119"
---
# <a name="context"></a><span data-ttu-id="2d713-103">context</span><span class="sxs-lookup"><span data-stu-id="2d713-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="2d713-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="2d713-104">[Office](office.md).context</span></span>

<span data-ttu-id="2d713-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="2d713-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="2d713-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.3).</span><span class="sxs-lookup"><span data-stu-id="2d713-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d713-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-107">Requirements</span></span>

|<span data-ttu-id="2d713-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-108">Requirement</span></span>| <span data-ttu-id="2d713-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-111">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-111">1.1</span></span>|
|[<span data-ttu-id="2d713-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="2d713-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="2d713-114">Properties</span></span>

| <span data-ttu-id="2d713-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="2d713-115">Property</span></span> | <span data-ttu-id="2d713-116">Modes</span><span class="sxs-lookup"><span data-stu-id="2d713-116">Modes</span></span> | <span data-ttu-id="2d713-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="2d713-117">Return type</span></span> | <span data-ttu-id="2d713-118">Minimale</span><span class="sxs-lookup"><span data-stu-id="2d713-118">Minimum</span></span><br><span data-ttu-id="2d713-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="2d713-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="2d713-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="2d713-121">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-121">Compose</span></span><br><span data-ttu-id="2d713-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-122">Read</span></span> | <span data-ttu-id="2d713-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2d713-123">String</span></span> | [<span data-ttu-id="2d713-124">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="2d713-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="2d713-126">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-126">Compose</span></span><br><span data-ttu-id="2d713-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-127">Read</span></span> | [<span data-ttu-id="2d713-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="2d713-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="2d713-129">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="2d713-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="2d713-131">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-131">Compose</span></span><br><span data-ttu-id="2d713-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-132">Read</span></span> | <span data-ttu-id="2d713-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="2d713-133">String</span></span> | [<span data-ttu-id="2d713-134">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-135">hote</span><span class="sxs-lookup"><span data-stu-id="2d713-135">host</span></span>](#host-hosttype) | <span data-ttu-id="2d713-136">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-136">Compose</span></span><br><span data-ttu-id="2d713-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-137">Read</span></span> | [<span data-ttu-id="2d713-138">HostType</span><span class="sxs-lookup"><span data-stu-id="2d713-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="2d713-139">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="2d713-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="2d713-141">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-141">Compose</span></span><br><span data-ttu-id="2d713-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-142">Read</span></span> | [<span data-ttu-id="2d713-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="2d713-144">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-145">plateforme</span><span class="sxs-lookup"><span data-stu-id="2d713-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="2d713-146">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-146">Compose</span></span><br><span data-ttu-id="2d713-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-147">Read</span></span> | [<span data-ttu-id="2d713-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="2d713-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="2d713-149">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-150">requise</span><span class="sxs-lookup"><span data-stu-id="2d713-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="2d713-151">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-151">Compose</span></span><br><span data-ttu-id="2d713-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-152">Read</span></span> | [<span data-ttu-id="2d713-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="2d713-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="2d713-154">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="2d713-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="2d713-156">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-156">Compose</span></span><br><span data-ttu-id="2d713-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-157">Read</span></span> | [<span data-ttu-id="2d713-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2d713-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="2d713-159">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="2d713-160">ui</span><span class="sxs-lookup"><span data-stu-id="2d713-160">ui</span></span>](#ui-ui) | <span data-ttu-id="2d713-161">Composition</span><span class="sxs-lookup"><span data-stu-id="2d713-161">Compose</span></span><br><span data-ttu-id="2d713-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-162">Read</span></span> | [<span data-ttu-id="2d713-163">UI</span><span class="sxs-lookup"><span data-stu-id="2d713-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="2d713-164">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="2d713-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="2d713-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="2d713-166">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="2d713-166">contentLanguage: String</span></span>

<span data-ttu-id="2d713-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="2d713-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="2d713-168">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="2d713-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-169">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-169">Type</span></span>

*   <span data-ttu-id="2d713-170">String</span><span class="sxs-lookup"><span data-stu-id="2d713-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d713-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-171">Requirements</span></span>

|<span data-ttu-id="2d713-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-172">Requirement</span></span>| <span data-ttu-id="2d713-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-175">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-175">1.1</span></span>|
|[<span data-ttu-id="2d713-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d713-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="2d713-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="2d713-179">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="2d713-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="2d713-180">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="2d713-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-181">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-181">Type</span></span>

*   [<span data-ttu-id="2d713-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="2d713-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="2d713-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-183">Requirements</span></span>

|<span data-ttu-id="2d713-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-184">Requirement</span></span>| <span data-ttu-id="2d713-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-187">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-187">1.1</span></span>|
|[<span data-ttu-id="2d713-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d713-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="2d713-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="2d713-191">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="2d713-191">displayLanguage: String</span></span>

<span data-ttu-id="2d713-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="2d713-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="2d713-193">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="2d713-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-194">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-194">Type</span></span>

*   <span data-ttu-id="2d713-195">String</span><span class="sxs-lookup"><span data-stu-id="2d713-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d713-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-196">Requirements</span></span>

|<span data-ttu-id="2d713-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-197">Requirement</span></span>| <span data-ttu-id="2d713-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-200">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-200">1.1</span></span>|
|[<span data-ttu-id="2d713-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d713-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="2d713-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="2d713-204">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="2d713-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="2d713-205">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="2d713-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-206">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-206">Type</span></span>

*   [<span data-ttu-id="2d713-207">HostType</span><span class="sxs-lookup"><span data-stu-id="2d713-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="2d713-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-208">Requirements</span></span>

|<span data-ttu-id="2d713-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-209">Requirement</span></span>| <span data-ttu-id="2d713-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-212">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-212">1.1</span></span>|
|[<span data-ttu-id="2d713-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-214">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d713-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="2d713-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="2d713-216">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="2d713-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="2d713-217">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="2d713-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-218">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-218">Type</span></span>

*   [<span data-ttu-id="2d713-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="2d713-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="2d713-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-220">Requirements</span></span>

|<span data-ttu-id="2d713-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-221">Requirement</span></span>| <span data-ttu-id="2d713-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-224">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-224">1.1</span></span>|
|[<span data-ttu-id="2d713-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d713-227">Exemple</span><span class="sxs-lookup"><span data-stu-id="2d713-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="2d713-228">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="2d713-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="2d713-229">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="2d713-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-230">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-230">Type</span></span>

*   [<span data-ttu-id="2d713-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="2d713-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="2d713-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-232">Requirements</span></span>

|<span data-ttu-id="2d713-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-233">Requirement</span></span>| <span data-ttu-id="2d713-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-236">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-236">1.1</span></span>|
|[<span data-ttu-id="2d713-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d713-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="2d713-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="2d713-240">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="2d713-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="2d713-241">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="2d713-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="2d713-242">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="2d713-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-243">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-243">Type</span></span>

*   [<span data-ttu-id="2d713-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2d713-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="2d713-245">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-245">Requirements</span></span>

|<span data-ttu-id="2d713-246">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-246">Requirement</span></span>| <span data-ttu-id="2d713-247">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-248">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-249">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-249">1.1</span></span>|
|[<span data-ttu-id="2d713-250">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="2d713-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="2d713-251">Restreinte</span><span class="sxs-lookup"><span data-stu-id="2d713-251">Restricted</span></span>|
|[<span data-ttu-id="2d713-252">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-253">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="2d713-254">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="2d713-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="2d713-255">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="2d713-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="2d713-256">Type</span><span class="sxs-lookup"><span data-stu-id="2d713-256">Type</span></span>

*   [<span data-ttu-id="2d713-257">UI</span><span class="sxs-lookup"><span data-stu-id="2d713-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="2d713-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="2d713-258">Requirements</span></span>

|<span data-ttu-id="2d713-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="2d713-259">Requirement</span></span>| <span data-ttu-id="2d713-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="2d713-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d713-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="2d713-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="2d713-262">1.1</span><span class="sxs-lookup"><span data-stu-id="2d713-262">1.1</span></span>|
|[<span data-ttu-id="2d713-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="2d713-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="2d713-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="2d713-264">Compose or Read</span></span>|
