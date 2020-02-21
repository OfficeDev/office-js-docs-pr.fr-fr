---
title: Office. Context-ensemble de conditions requises 1,6
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: ceca9e717f84a596e98ef9d6aade906323ef09fb
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165369"
---
# <a name="context"></a><span data-ttu-id="28d01-102">context</span><span class="sxs-lookup"><span data-stu-id="28d01-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="28d01-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="28d01-103">[Office](office.md).context</span></span>

<span data-ttu-id="28d01-104">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="28d01-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="28d01-105">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.6).</span><span class="sxs-lookup"><span data-stu-id="28d01-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6).</span></span>

##### <a name="requirements"></a><span data-ttu-id="28d01-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-106">Requirements</span></span>

|<span data-ttu-id="28d01-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-107">Requirement</span></span>| <span data-ttu-id="28d01-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-110">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-110">1.1</span></span>|
|[<span data-ttu-id="28d01-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="28d01-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="28d01-113">Properties</span></span>

| <span data-ttu-id="28d01-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="28d01-114">Property</span></span> | <span data-ttu-id="28d01-115">Modes</span><span class="sxs-lookup"><span data-stu-id="28d01-115">Modes</span></span> | <span data-ttu-id="28d01-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="28d01-116">Return type</span></span> | <span data-ttu-id="28d01-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="28d01-117">Minimum</span></span><br><span data-ttu-id="28d01-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="28d01-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="28d01-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="28d01-120">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-120">Compose</span></span><br><span data-ttu-id="28d01-121">Lire</span><span class="sxs-lookup"><span data-stu-id="28d01-121">Read</span></span> | <span data-ttu-id="28d01-122">Chaîne</span><span class="sxs-lookup"><span data-stu-id="28d01-122">String</span></span> | [<span data-ttu-id="28d01-123">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-124">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="28d01-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="28d01-125">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-125">Compose</span></span><br><span data-ttu-id="28d01-126">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-126">Read</span></span> | [<span data-ttu-id="28d01-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="28d01-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6) | [<span data-ttu-id="28d01-128">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="28d01-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="28d01-130">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-130">Compose</span></span><br><span data-ttu-id="28d01-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-131">Read</span></span> | <span data-ttu-id="28d01-132">Chaîne</span><span class="sxs-lookup"><span data-stu-id="28d01-132">String</span></span> | [<span data-ttu-id="28d01-133">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-134">hote</span><span class="sxs-lookup"><span data-stu-id="28d01-134">host</span></span>](#host-hosttype) | <span data-ttu-id="28d01-135">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-135">Compose</span></span><br><span data-ttu-id="28d01-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-136">Read</span></span> | [<span data-ttu-id="28d01-137">HostType</span><span class="sxs-lookup"><span data-stu-id="28d01-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6) | [<span data-ttu-id="28d01-138">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="28d01-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="28d01-140">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-140">Compose</span></span><br><span data-ttu-id="28d01-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-141">Read</span></span> | [<span data-ttu-id="28d01-142">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6) | [<span data-ttu-id="28d01-143">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-144">plateforme</span><span class="sxs-lookup"><span data-stu-id="28d01-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="28d01-145">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-145">Compose</span></span><br><span data-ttu-id="28d01-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-146">Read</span></span> | [<span data-ttu-id="28d01-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="28d01-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6) | [<span data-ttu-id="28d01-148">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-149">requise</span><span class="sxs-lookup"><span data-stu-id="28d01-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="28d01-150">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-150">Compose</span></span><br><span data-ttu-id="28d01-151">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-151">Read</span></span> | [<span data-ttu-id="28d01-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="28d01-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6) | [<span data-ttu-id="28d01-153">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="28d01-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="28d01-155">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-155">Compose</span></span><br><span data-ttu-id="28d01-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-156">Read</span></span> | [<span data-ttu-id="28d01-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="28d01-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6) | [<span data-ttu-id="28d01-158">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="28d01-159">ui</span><span class="sxs-lookup"><span data-stu-id="28d01-159">ui</span></span>](#ui-ui) | <span data-ttu-id="28d01-160">Composition</span><span class="sxs-lookup"><span data-stu-id="28d01-160">Compose</span></span><br><span data-ttu-id="28d01-161">Lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-161">Read</span></span> | [<span data-ttu-id="28d01-162">UI</span><span class="sxs-lookup"><span data-stu-id="28d01-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6) | [<span data-ttu-id="28d01-163">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="28d01-164">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="28d01-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="28d01-165">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="28d01-165">contentLanguage: String</span></span>

<span data-ttu-id="28d01-166">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="28d01-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="28d01-167">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="28d01-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-168">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-168">Type</span></span>

*   <span data-ttu-id="28d01-169">String</span><span class="sxs-lookup"><span data-stu-id="28d01-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28d01-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-170">Requirements</span></span>

|<span data-ttu-id="28d01-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-171">Requirement</span></span>| <span data-ttu-id="28d01-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-174">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-174">1.1</span></span>|
|[<span data-ttu-id="28d01-175">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-175">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-176">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28d01-177">Exemple</span><span class="sxs-lookup"><span data-stu-id="28d01-177">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="28d01-178">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="28d01-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="28d01-179">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="28d01-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-180">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-180">Type</span></span>

*   [<span data-ttu-id="28d01-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="28d01-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="28d01-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-182">Requirements</span></span>

|<span data-ttu-id="28d01-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-183">Requirement</span></span>| <span data-ttu-id="28d01-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-186">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-186">1.1</span></span>|
|[<span data-ttu-id="28d01-187">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-187">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-188">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28d01-189">Exemple</span><span class="sxs-lookup"><span data-stu-id="28d01-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="28d01-190">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="28d01-190">displayLanguage: String</span></span>

<span data-ttu-id="28d01-191">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="28d01-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="28d01-192">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="28d01-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-193">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-193">Type</span></span>

*   <span data-ttu-id="28d01-194">String</span><span class="sxs-lookup"><span data-stu-id="28d01-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="28d01-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-195">Requirements</span></span>

|<span data-ttu-id="28d01-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-196">Requirement</span></span>| <span data-ttu-id="28d01-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-199">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-199">1.1</span></span>|
|[<span data-ttu-id="28d01-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-200">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-201">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28d01-202">Exemple</span><span class="sxs-lookup"><span data-stu-id="28d01-202">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="28d01-203">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="28d01-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="28d01-204">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="28d01-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-205">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-205">Type</span></span>

*   [<span data-ttu-id="28d01-206">HostType</span><span class="sxs-lookup"><span data-stu-id="28d01-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="28d01-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-207">Requirements</span></span>

|<span data-ttu-id="28d01-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-208">Requirement</span></span>| <span data-ttu-id="28d01-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-211">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-211">1.1</span></span>|
|[<span data-ttu-id="28d01-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-212">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-213">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28d01-214">Exemple</span><span class="sxs-lookup"><span data-stu-id="28d01-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="28d01-215">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="28d01-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="28d01-216">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="28d01-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-217">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-217">Type</span></span>

*   [<span data-ttu-id="28d01-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="28d01-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="28d01-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-219">Requirements</span></span>

|<span data-ttu-id="28d01-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-220">Requirement</span></span>| <span data-ttu-id="28d01-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-223">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-223">1.1</span></span>|
|[<span data-ttu-id="28d01-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28d01-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="28d01-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="28d01-227">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="28d01-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="28d01-228">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="28d01-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-229">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-229">Type</span></span>

*   [<span data-ttu-id="28d01-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="28d01-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="28d01-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-231">Requirements</span></span>

|<span data-ttu-id="28d01-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-232">Requirement</span></span>| <span data-ttu-id="28d01-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-235">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-235">1.1</span></span>|
|[<span data-ttu-id="28d01-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="28d01-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="28d01-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="28d01-239">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="28d01-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="28d01-240">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="28d01-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="28d01-241">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="28d01-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-242">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-242">Type</span></span>

*   [<span data-ttu-id="28d01-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="28d01-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="28d01-244">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-244">Requirements</span></span>

|<span data-ttu-id="28d01-245">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-245">Requirement</span></span>| <span data-ttu-id="28d01-246">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-247">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-248">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-248">1.1</span></span>|
|[<span data-ttu-id="28d01-249">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="28d01-249">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="28d01-250">Restreinte</span><span class="sxs-lookup"><span data-stu-id="28d01-250">Restricted</span></span>|
|[<span data-ttu-id="28d01-251">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-251">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-252">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="28d01-253">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="28d01-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="28d01-254">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="28d01-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="28d01-255">Type</span><span class="sxs-lookup"><span data-stu-id="28d01-255">Type</span></span>

*   [<span data-ttu-id="28d01-256">UI</span><span class="sxs-lookup"><span data-stu-id="28d01-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="28d01-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="28d01-257">Requirements</span></span>

|<span data-ttu-id="28d01-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="28d01-258">Requirement</span></span>| <span data-ttu-id="28d01-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="28d01-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="28d01-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="28d01-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="28d01-261">1.1</span><span class="sxs-lookup"><span data-stu-id="28d01-261">1.1</span></span>|
|[<span data-ttu-id="28d01-262">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="28d01-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="28d01-263">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="28d01-263">Compose or Read</span></span>|
