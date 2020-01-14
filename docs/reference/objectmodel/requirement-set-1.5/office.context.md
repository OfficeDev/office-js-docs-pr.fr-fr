---
title: Office. Context-ensemble de conditions requises 1,5
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: c6e7c3f9b37942454843ed9f501f80a5f4af6e55
ms.sourcegitcommit: 0dacbe7c80ed387099e3ec21e151f8990b181ede
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/13/2020
ms.locfileid: "41110954"
---
# <a name="context"></a><span data-ttu-id="146b4-102">context</span><span class="sxs-lookup"><span data-stu-id="146b4-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="146b4-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="146b4-103">[Office](office.md).context</span></span>

<span data-ttu-id="146b4-104">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="146b4-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="146b4-105">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.5).</span><span class="sxs-lookup"><span data-stu-id="146b4-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5).</span></span>

##### <a name="requirements"></a><span data-ttu-id="146b4-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-106">Requirements</span></span>

|<span data-ttu-id="146b4-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-107">Requirement</span></span>| <span data-ttu-id="146b4-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-110">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-110">1.1</span></span>|
|[<span data-ttu-id="146b4-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="146b4-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="146b4-113">Properties</span></span>

| <span data-ttu-id="146b4-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="146b4-114">Property</span></span> | <span data-ttu-id="146b4-115">Modes</span><span class="sxs-lookup"><span data-stu-id="146b4-115">Modes</span></span> | <span data-ttu-id="146b4-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="146b4-116">Return type</span></span> | <span data-ttu-id="146b4-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="146b4-117">Minimum</span></span><br><span data-ttu-id="146b4-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="146b4-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="146b4-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="146b4-120">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-120">Compose</span></span><br><span data-ttu-id="146b4-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-121">Read</span></span> | <span data-ttu-id="146b4-122">Chaîne</span><span class="sxs-lookup"><span data-stu-id="146b4-122">String</span></span> | [<span data-ttu-id="146b4-123">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-124">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="146b4-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="146b4-125">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-125">Compose</span></span><br><span data-ttu-id="146b4-126">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-126">Read</span></span> | [<span data-ttu-id="146b4-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="146b4-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5) | [<span data-ttu-id="146b4-128">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="146b4-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="146b4-130">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-130">Compose</span></span><br><span data-ttu-id="146b4-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-131">Read</span></span> | <span data-ttu-id="146b4-132">Chaîne</span><span class="sxs-lookup"><span data-stu-id="146b4-132">String</span></span> | [<span data-ttu-id="146b4-133">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-134">hote</span><span class="sxs-lookup"><span data-stu-id="146b4-134">host</span></span>](#host-hosttype) | <span data-ttu-id="146b4-135">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-135">Compose</span></span><br><span data-ttu-id="146b4-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-136">Read</span></span> | [<span data-ttu-id="146b4-137">HostType</span><span class="sxs-lookup"><span data-stu-id="146b4-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5) | [<span data-ttu-id="146b4-138">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="146b4-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="146b4-140">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-140">Compose</span></span><br><span data-ttu-id="146b4-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-141">Read</span></span> | [<span data-ttu-id="146b4-142">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5) | [<span data-ttu-id="146b4-143">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-144">plateforme</span><span class="sxs-lookup"><span data-stu-id="146b4-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="146b4-145">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-145">Compose</span></span><br><span data-ttu-id="146b4-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-146">Read</span></span> | [<span data-ttu-id="146b4-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="146b4-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5) | [<span data-ttu-id="146b4-148">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-149">requise</span><span class="sxs-lookup"><span data-stu-id="146b4-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="146b4-150">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-150">Compose</span></span><br><span data-ttu-id="146b4-151">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-151">Read</span></span> | [<span data-ttu-id="146b4-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="146b4-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5) | [<span data-ttu-id="146b4-153">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="146b4-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="146b4-155">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-155">Compose</span></span><br><span data-ttu-id="146b4-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-156">Read</span></span> | [<span data-ttu-id="146b4-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="146b4-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5) | [<span data-ttu-id="146b4-158">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="146b4-159">ui</span><span class="sxs-lookup"><span data-stu-id="146b4-159">ui</span></span>](#ui-ui) | <span data-ttu-id="146b4-160">Composition</span><span class="sxs-lookup"><span data-stu-id="146b4-160">Compose</span></span><br><span data-ttu-id="146b4-161">Lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-161">Read</span></span> | [<span data-ttu-id="146b4-162">UI</span><span class="sxs-lookup"><span data-stu-id="146b4-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5) | [<span data-ttu-id="146b4-163">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="146b4-164">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="146b4-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="146b4-165">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="146b4-165">contentLanguage: String</span></span>

<span data-ttu-id="146b4-166">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="146b4-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="146b4-167">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="146b4-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-168">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-168">Type</span></span>

*   <span data-ttu-id="146b4-169">String</span><span class="sxs-lookup"><span data-stu-id="146b4-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="146b4-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-170">Requirements</span></span>

|<span data-ttu-id="146b4-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-171">Requirement</span></span>| <span data-ttu-id="146b4-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-174">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-174">1.1</span></span>|
|[<span data-ttu-id="146b4-175">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-176">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="146b4-177">Exemple</span><span class="sxs-lookup"><span data-stu-id="146b4-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="146b4-178">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="146b4-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="146b4-179">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="146b4-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-180">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-180">Type</span></span>

*   [<span data-ttu-id="146b4-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="146b4-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="146b4-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-182">Requirements</span></span>

|<span data-ttu-id="146b4-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-183">Requirement</span></span>| <span data-ttu-id="146b4-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-186">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-186">1.1</span></span>|
|[<span data-ttu-id="146b4-187">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-188">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="146b4-189">Exemple</span><span class="sxs-lookup"><span data-stu-id="146b4-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="146b4-190">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="146b4-190">displayLanguage: String</span></span>

<span data-ttu-id="146b4-191">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="146b4-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="146b4-192">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="146b4-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-193">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-193">Type</span></span>

*   <span data-ttu-id="146b4-194">String</span><span class="sxs-lookup"><span data-stu-id="146b4-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="146b4-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-195">Requirements</span></span>

|<span data-ttu-id="146b4-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-196">Requirement</span></span>| <span data-ttu-id="146b4-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-199">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-199">1.1</span></span>|
|[<span data-ttu-id="146b4-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-201">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="146b4-202">Exemple</span><span class="sxs-lookup"><span data-stu-id="146b4-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="146b4-203">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="146b4-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="146b4-204">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="146b4-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-205">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-205">Type</span></span>

*   [<span data-ttu-id="146b4-206">HostType</span><span class="sxs-lookup"><span data-stu-id="146b4-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="146b4-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-207">Requirements</span></span>

|<span data-ttu-id="146b4-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-208">Requirement</span></span>| <span data-ttu-id="146b4-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-211">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-211">1.1</span></span>|
|[<span data-ttu-id="146b4-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-213">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="146b4-214">Exemple</span><span class="sxs-lookup"><span data-stu-id="146b4-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="146b4-215">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="146b4-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="146b4-216">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="146b4-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-217">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-217">Type</span></span>

*   [<span data-ttu-id="146b4-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="146b4-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="146b4-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-219">Requirements</span></span>

|<span data-ttu-id="146b4-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-220">Requirement</span></span>| <span data-ttu-id="146b4-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-223">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-223">1.1</span></span>|
|[<span data-ttu-id="146b4-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="146b4-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="146b4-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="146b4-227">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="146b4-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="146b4-228">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="146b4-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-229">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-229">Type</span></span>

*   [<span data-ttu-id="146b4-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="146b4-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="146b4-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-231">Requirements</span></span>

|<span data-ttu-id="146b4-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-232">Requirement</span></span>| <span data-ttu-id="146b4-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-235">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-235">1.1</span></span>|
|[<span data-ttu-id="146b4-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="146b4-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="146b4-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="146b4-239">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="146b4-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="146b4-240">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="146b4-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="146b4-241">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="146b4-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-242">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-242">Type</span></span>

*   [<span data-ttu-id="146b4-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="146b4-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="146b4-244">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-244">Requirements</span></span>

|<span data-ttu-id="146b4-245">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-245">Requirement</span></span>| <span data-ttu-id="146b4-246">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-247">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-248">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-248">1.1</span></span>|
|[<span data-ttu-id="146b4-249">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="146b4-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="146b4-250">Restreinte</span><span class="sxs-lookup"><span data-stu-id="146b4-250">Restricted</span></span>|
|[<span data-ttu-id="146b4-251">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-252">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="146b4-253">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="146b4-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="146b4-254">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="146b4-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="146b4-255">Type</span><span class="sxs-lookup"><span data-stu-id="146b4-255">Type</span></span>

*   [<span data-ttu-id="146b4-256">UI</span><span class="sxs-lookup"><span data-stu-id="146b4-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="146b4-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="146b4-257">Requirements</span></span>

|<span data-ttu-id="146b4-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="146b4-258">Requirement</span></span>| <span data-ttu-id="146b4-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="146b4-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="146b4-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="146b4-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="146b4-261">1.1</span><span class="sxs-lookup"><span data-stu-id="146b4-261">1.1</span></span>|
|[<span data-ttu-id="146b4-262">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="146b4-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="146b4-263">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="146b4-263">Compose or Read</span></span>|
