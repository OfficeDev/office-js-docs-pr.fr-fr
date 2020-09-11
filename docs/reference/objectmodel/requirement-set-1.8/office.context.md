---
title: Office. Context-ensemble de conditions requises 1,8
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,8.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 755ef23633182fdedb1fe199118b985a8319e4ad
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431527"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="91893-103">contexte (boîte aux lettres requise définie sur 1,8)</span><span class="sxs-lookup"><span data-stu-id="91893-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="91893-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="91893-104">[Office](office.md).context</span></span>

<span data-ttu-id="91893-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="91893-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="91893-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="91893-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="91893-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-107">Requirements</span></span>

|<span data-ttu-id="91893-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-108">Requirement</span></span>| <span data-ttu-id="91893-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-111">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-111">1.1</span></span>|
|[<span data-ttu-id="91893-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="91893-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="91893-114">Properties</span></span>

| <span data-ttu-id="91893-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="91893-115">Property</span></span> | <span data-ttu-id="91893-116">Modes</span><span class="sxs-lookup"><span data-stu-id="91893-116">Modes</span></span> | <span data-ttu-id="91893-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="91893-117">Return type</span></span> | <span data-ttu-id="91893-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="91893-118">Minimum</span></span><br><span data-ttu-id="91893-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="91893-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="91893-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="91893-121">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-121">Compose</span></span><br><span data-ttu-id="91893-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-122">Read</span></span> | <span data-ttu-id="91893-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="91893-123">String</span></span> | [<span data-ttu-id="91893-124">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="91893-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="91893-126">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-126">Compose</span></span><br><span data-ttu-id="91893-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-127">Read</span></span> | [<span data-ttu-id="91893-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="91893-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="91893-129">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="91893-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="91893-131">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-131">Compose</span></span><br><span data-ttu-id="91893-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-132">Read</span></span> | <span data-ttu-id="91893-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="91893-133">String</span></span> | [<span data-ttu-id="91893-134">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-135">hote</span><span class="sxs-lookup"><span data-stu-id="91893-135">host</span></span>](#host-hosttype) | <span data-ttu-id="91893-136">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-136">Compose</span></span><br><span data-ttu-id="91893-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-137">Read</span></span> | [<span data-ttu-id="91893-138">HostType</span><span class="sxs-lookup"><span data-stu-id="91893-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="91893-139">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="91893-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="91893-141">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-141">Compose</span></span><br><span data-ttu-id="91893-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-142">Read</span></span> | [<span data-ttu-id="91893-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="91893-144">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-145">plateforme</span><span class="sxs-lookup"><span data-stu-id="91893-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="91893-146">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-146">Compose</span></span><br><span data-ttu-id="91893-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-147">Read</span></span> | [<span data-ttu-id="91893-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="91893-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="91893-149">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-150">requise</span><span class="sxs-lookup"><span data-stu-id="91893-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="91893-151">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-151">Compose</span></span><br><span data-ttu-id="91893-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-152">Read</span></span> | [<span data-ttu-id="91893-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="91893-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="91893-154">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="91893-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="91893-156">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-156">Compose</span></span><br><span data-ttu-id="91893-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-157">Read</span></span> | [<span data-ttu-id="91893-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="91893-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="91893-159">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="91893-160">ui</span><span class="sxs-lookup"><span data-stu-id="91893-160">ui</span></span>](#ui-ui) | <span data-ttu-id="91893-161">Composition</span><span class="sxs-lookup"><span data-stu-id="91893-161">Compose</span></span><br><span data-ttu-id="91893-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="91893-162">Read</span></span> | [<span data-ttu-id="91893-163">UI</span><span class="sxs-lookup"><span data-stu-id="91893-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="91893-164">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="91893-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="91893-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="91893-166">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="91893-166">contentLanguage: String</span></span>

<span data-ttu-id="91893-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="91893-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="91893-168">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="91893-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-169">Type</span><span class="sxs-lookup"><span data-stu-id="91893-169">Type</span></span>

*   <span data-ttu-id="91893-170">String</span><span class="sxs-lookup"><span data-stu-id="91893-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91893-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-171">Requirements</span></span>

|<span data-ttu-id="91893-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-172">Requirement</span></span>| <span data-ttu-id="91893-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-175">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-175">1.1</span></span>|
|[<span data-ttu-id="91893-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91893-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="91893-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="91893-179">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="91893-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="91893-180">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="91893-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-181">Type</span><span class="sxs-lookup"><span data-stu-id="91893-181">Type</span></span>

*   [<span data-ttu-id="91893-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="91893-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="91893-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-183">Requirements</span></span>

|<span data-ttu-id="91893-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-184">Requirement</span></span>| <span data-ttu-id="91893-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-187">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-187">1.1</span></span>|
|[<span data-ttu-id="91893-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91893-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="91893-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="91893-191">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="91893-191">displayLanguage: String</span></span>

<span data-ttu-id="91893-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="91893-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="91893-193">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="91893-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-194">Type</span><span class="sxs-lookup"><span data-stu-id="91893-194">Type</span></span>

*   <span data-ttu-id="91893-195">String</span><span class="sxs-lookup"><span data-stu-id="91893-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91893-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-196">Requirements</span></span>

|<span data-ttu-id="91893-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-197">Requirement</span></span>| <span data-ttu-id="91893-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-200">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-200">1.1</span></span>|
|[<span data-ttu-id="91893-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91893-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="91893-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="91893-204">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="91893-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="91893-205">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="91893-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-206">Type</span><span class="sxs-lookup"><span data-stu-id="91893-206">Type</span></span>

*   [<span data-ttu-id="91893-207">HostType</span><span class="sxs-lookup"><span data-stu-id="91893-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="91893-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-208">Requirements</span></span>

|<span data-ttu-id="91893-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-209">Requirement</span></span>| <span data-ttu-id="91893-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-212">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-212">1.1</span></span>|
|[<span data-ttu-id="91893-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-214">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91893-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="91893-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="91893-216">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="91893-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="91893-217">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="91893-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-218">Type</span><span class="sxs-lookup"><span data-stu-id="91893-218">Type</span></span>

*   [<span data-ttu-id="91893-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="91893-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="91893-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-220">Requirements</span></span>

|<span data-ttu-id="91893-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-221">Requirement</span></span>| <span data-ttu-id="91893-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-224">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-224">1.1</span></span>|
|[<span data-ttu-id="91893-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91893-227">Exemple</span><span class="sxs-lookup"><span data-stu-id="91893-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="91893-228">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="91893-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="91893-229">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="91893-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-230">Type</span><span class="sxs-lookup"><span data-stu-id="91893-230">Type</span></span>

*   [<span data-ttu-id="91893-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="91893-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="91893-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-232">Requirements</span></span>

|<span data-ttu-id="91893-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-233">Requirement</span></span>| <span data-ttu-id="91893-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-236">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-236">1.1</span></span>|
|[<span data-ttu-id="91893-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91893-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="91893-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="91893-240">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="91893-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="91893-241">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="91893-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="91893-242">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="91893-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-243">Type</span><span class="sxs-lookup"><span data-stu-id="91893-243">Type</span></span>

*   [<span data-ttu-id="91893-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="91893-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="91893-245">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-245">Requirements</span></span>

|<span data-ttu-id="91893-246">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-246">Requirement</span></span>| <span data-ttu-id="91893-247">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-248">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-249">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-249">1.1</span></span>|
|[<span data-ttu-id="91893-250">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="91893-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="91893-251">Restreinte</span><span class="sxs-lookup"><span data-stu-id="91893-251">Restricted</span></span>|
|[<span data-ttu-id="91893-252">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-253">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="91893-254">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="91893-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="91893-255">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="91893-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="91893-256">Type</span><span class="sxs-lookup"><span data-stu-id="91893-256">Type</span></span>

*   [<span data-ttu-id="91893-257">UI</span><span class="sxs-lookup"><span data-stu-id="91893-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="91893-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="91893-258">Requirements</span></span>

|<span data-ttu-id="91893-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="91893-259">Requirement</span></span>| <span data-ttu-id="91893-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="91893-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="91893-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="91893-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="91893-262">1.1</span><span class="sxs-lookup"><span data-stu-id="91893-262">1.1</span></span>|
|[<span data-ttu-id="91893-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="91893-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="91893-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="91893-264">Compose or Read</span></span>|
