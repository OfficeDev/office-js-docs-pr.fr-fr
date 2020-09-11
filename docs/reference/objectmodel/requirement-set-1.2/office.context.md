---
title: Office.context-ensemble de conditions requises 1.2
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,2.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: da9df156ad75bf437a12812fe9144e0fd6a9586f
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431289"
---
# <a name="context-mailbox-requirement-set-12"></a><span data-ttu-id="e4a85-103">contexte (boîte aux lettres requise définie sur 1,2)</span><span class="sxs-lookup"><span data-stu-id="e4a85-103">context (Mailbox requirement set 1.2)</span></span>

### <a name="officecontext"></a><span data-ttu-id="e4a85-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="e4a85-104">[Office](office.md).context</span></span>

<span data-ttu-id="e4a85-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="e4a85-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="e4a85-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="e4a85-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e4a85-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-107">Requirements</span></span>

|<span data-ttu-id="e4a85-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-108">Requirement</span></span>| <span data-ttu-id="e4a85-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-111">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-111">1.1</span></span>|
|[<span data-ttu-id="e4a85-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e4a85-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e4a85-114">Properties</span></span>

| <span data-ttu-id="e4a85-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="e4a85-115">Property</span></span> | <span data-ttu-id="e4a85-116">Modes</span><span class="sxs-lookup"><span data-stu-id="e4a85-116">Modes</span></span> | <span data-ttu-id="e4a85-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="e4a85-117">Return type</span></span> | <span data-ttu-id="e4a85-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="e4a85-118">Minimum</span></span><br><span data-ttu-id="e4a85-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e4a85-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="e4a85-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="e4a85-121">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-121">Compose</span></span><br><span data-ttu-id="e4a85-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-122">Read</span></span> | <span data-ttu-id="e4a85-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e4a85-123">String</span></span> | [<span data-ttu-id="e4a85-124">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="e4a85-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="e4a85-126">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-126">Compose</span></span><br><span data-ttu-id="e4a85-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-127">Read</span></span> | [<span data-ttu-id="e4a85-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e4a85-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="e4a85-129">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="e4a85-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="e4a85-131">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-131">Compose</span></span><br><span data-ttu-id="e4a85-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-132">Read</span></span> | <span data-ttu-id="e4a85-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e4a85-133">String</span></span> | [<span data-ttu-id="e4a85-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-135">hote</span><span class="sxs-lookup"><span data-stu-id="e4a85-135">host</span></span>](#host-hosttype) | <span data-ttu-id="e4a85-136">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-136">Compose</span></span><br><span data-ttu-id="e4a85-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-137">Read</span></span> | [<span data-ttu-id="e4a85-138">HostType</span><span class="sxs-lookup"><span data-stu-id="e4a85-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="e4a85-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="e4a85-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="e4a85-141">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-141">Compose</span></span><br><span data-ttu-id="e4a85-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-142">Read</span></span> | [<span data-ttu-id="e4a85-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="e4a85-144">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-145">plateforme</span><span class="sxs-lookup"><span data-stu-id="e4a85-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="e4a85-146">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-146">Compose</span></span><br><span data-ttu-id="e4a85-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-147">Read</span></span> | [<span data-ttu-id="e4a85-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e4a85-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="e4a85-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-150">requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="e4a85-151">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-151">Compose</span></span><br><span data-ttu-id="e4a85-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-152">Read</span></span> | [<span data-ttu-id="e4a85-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e4a85-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="e4a85-154">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="e4a85-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="e4a85-156">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-156">Compose</span></span><br><span data-ttu-id="e4a85-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-157">Read</span></span> | [<span data-ttu-id="e4a85-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e4a85-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="e4a85-159">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e4a85-160">ui</span><span class="sxs-lookup"><span data-stu-id="e4a85-160">ui</span></span>](#ui-ui) | <span data-ttu-id="e4a85-161">Composition</span><span class="sxs-lookup"><span data-stu-id="e4a85-161">Compose</span></span><br><span data-ttu-id="e4a85-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-162">Read</span></span> | [<span data-ttu-id="e4a85-163">UI</span><span class="sxs-lookup"><span data-stu-id="e4a85-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="e4a85-164">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="e4a85-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="e4a85-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="e4a85-166">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="e4a85-166">contentLanguage: String</span></span>

<span data-ttu-id="e4a85-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="e4a85-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="e4a85-168">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="e4a85-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-169">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-169">Type</span></span>

*   <span data-ttu-id="e4a85-170">String</span><span class="sxs-lookup"><span data-stu-id="e4a85-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e4a85-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-171">Requirements</span></span>

|<span data-ttu-id="e4a85-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-172">Requirement</span></span>| <span data-ttu-id="e4a85-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-175">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-175">1.1</span></span>|
|[<span data-ttu-id="e4a85-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e4a85-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="e4a85-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="e4a85-179">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="e4a85-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="e4a85-180">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e4a85-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-181">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-181">Type</span></span>

*   [<span data-ttu-id="e4a85-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e4a85-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="e4a85-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-183">Requirements</span></span>

|<span data-ttu-id="e4a85-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-184">Requirement</span></span>| <span data-ttu-id="e4a85-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-187">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-187">1.1</span></span>|
|[<span data-ttu-id="e4a85-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e4a85-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="e4a85-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="e4a85-191">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="e4a85-191">displayLanguage: String</span></span>

<span data-ttu-id="e4a85-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="e4a85-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="e4a85-193">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="e4a85-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-194">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-194">Type</span></span>

*   <span data-ttu-id="e4a85-195">String</span><span class="sxs-lookup"><span data-stu-id="e4a85-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e4a85-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-196">Requirements</span></span>

|<span data-ttu-id="e4a85-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-197">Requirement</span></span>| <span data-ttu-id="e4a85-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-200">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-200">1.1</span></span>|
|[<span data-ttu-id="e4a85-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e4a85-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="e4a85-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="e4a85-204">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="e4a85-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="e4a85-205">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="e4a85-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-206">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-206">Type</span></span>

*   [<span data-ttu-id="e4a85-207">HostType</span><span class="sxs-lookup"><span data-stu-id="e4a85-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="e4a85-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-208">Requirements</span></span>

|<span data-ttu-id="e4a85-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-209">Requirement</span></span>| <span data-ttu-id="e4a85-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-212">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-212">1.1</span></span>|
|[<span data-ttu-id="e4a85-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-214">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e4a85-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="e4a85-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="e4a85-216">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="e4a85-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="e4a85-217">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e4a85-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-218">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-218">Type</span></span>

*   [<span data-ttu-id="e4a85-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e4a85-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="e4a85-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-220">Requirements</span></span>

|<span data-ttu-id="e4a85-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-221">Requirement</span></span>| <span data-ttu-id="e4a85-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-224">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-224">1.1</span></span>|
|[<span data-ttu-id="e4a85-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e4a85-227">Exemple</span><span class="sxs-lookup"><span data-stu-id="e4a85-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="e4a85-228">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="e4a85-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="e4a85-229">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="e4a85-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-230">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-230">Type</span></span>

*   [<span data-ttu-id="e4a85-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e4a85-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="e4a85-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-232">Requirements</span></span>

|<span data-ttu-id="e4a85-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-233">Requirement</span></span>| <span data-ttu-id="e4a85-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-236">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-236">1.1</span></span>|
|[<span data-ttu-id="e4a85-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e4a85-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="e4a85-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="e4a85-240">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="e4a85-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="e4a85-241">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e4a85-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e4a85-242">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e4a85-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-243">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-243">Type</span></span>

*   [<span data-ttu-id="e4a85-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e4a85-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e4a85-245">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-245">Requirements</span></span>

|<span data-ttu-id="e4a85-246">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-246">Requirement</span></span>| <span data-ttu-id="e4a85-247">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-248">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-249">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-249">1.1</span></span>|
|[<span data-ttu-id="e4a85-250">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e4a85-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="e4a85-251">Restreinte</span><span class="sxs-lookup"><span data-stu-id="e4a85-251">Restricted</span></span>|
|[<span data-ttu-id="e4a85-252">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-253">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="e4a85-254">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="e4a85-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="e4a85-255">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="e4a85-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e4a85-256">Type</span><span class="sxs-lookup"><span data-stu-id="e4a85-256">Type</span></span>

*   [<span data-ttu-id="e4a85-257">UI</span><span class="sxs-lookup"><span data-stu-id="e4a85-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="e4a85-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e4a85-258">Requirements</span></span>

|<span data-ttu-id="e4a85-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e4a85-259">Requirement</span></span>| <span data-ttu-id="e4a85-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="e4a85-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="e4a85-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e4a85-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e4a85-262">1.1</span><span class="sxs-lookup"><span data-stu-id="e4a85-262">1.1</span></span>|
|[<span data-ttu-id="e4a85-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e4a85-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e4a85-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e4a85-264">Compose or Read</span></span>|
