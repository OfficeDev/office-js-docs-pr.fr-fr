---
title: Office. Context-ensemble de conditions requises 1,8
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,8.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: cf49abb05bbe2e5e7b1d4d178c7749d6e7183d2a
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570764"
---
# <a name="context-mailbox-requirement-set-18"></a><span data-ttu-id="e83e5-103">contexte (boîte aux lettres requise définie sur 1,8)</span><span class="sxs-lookup"><span data-stu-id="e83e5-103">context (Mailbox requirement set 1.8)</span></span>

### <a name="officecontext"></a><span data-ttu-id="e83e5-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="e83e5-104">[Office](office.md).context</span></span>

<span data-ttu-id="e83e5-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="e83e5-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="e83e5-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="e83e5-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.8&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e83e5-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-107">Requirements</span></span>

|<span data-ttu-id="e83e5-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-108">Requirement</span></span>| <span data-ttu-id="e83e5-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-111">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-111">1.1</span></span>|
|[<span data-ttu-id="e83e5-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e83e5-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="e83e5-114">Properties</span></span>

| <span data-ttu-id="e83e5-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="e83e5-115">Property</span></span> | <span data-ttu-id="e83e5-116">Modes</span><span class="sxs-lookup"><span data-stu-id="e83e5-116">Modes</span></span> | <span data-ttu-id="e83e5-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="e83e5-117">Return type</span></span> | <span data-ttu-id="e83e5-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="e83e5-118">Minimum</span></span><br><span data-ttu-id="e83e5-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e83e5-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="e83e5-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="e83e5-121">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-121">Compose</span></span><br><span data-ttu-id="e83e5-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-122">Read</span></span> | <span data-ttu-id="e83e5-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e83e5-123">String</span></span> | [<span data-ttu-id="e83e5-124">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e83e5-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="e83e5-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="e83e5-126">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-126">Compose</span></span><br><span data-ttu-id="e83e5-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-127">Read</span></span> | [<span data-ttu-id="e83e5-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e83e5-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="e83e5-129">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e83e5-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="e83e5-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="e83e5-131">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-131">Compose</span></span><br><span data-ttu-id="e83e5-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-132">Read</span></span> | <span data-ttu-id="e83e5-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="e83e5-133">String</span></span> | [<span data-ttu-id="e83e5-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e83e5-135">hote</span><span class="sxs-lookup"><span data-stu-id="e83e5-135">host</span></span>](#host-hosttype) | <span data-ttu-id="e83e5-136">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-136">Compose</span></span><br><span data-ttu-id="e83e5-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-137">Read</span></span> | [<span data-ttu-id="e83e5-138">HostType</span><span class="sxs-lookup"><span data-stu-id="e83e5-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="e83e5-139">1,5</span><span class="sxs-lookup"><span data-stu-id="e83e5-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e83e5-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="e83e5-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="e83e5-141">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-141">Compose</span></span><br><span data-ttu-id="e83e5-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-142">Read</span></span> | [<span data-ttu-id="e83e5-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="e83e5-144">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e83e5-145">plateforme</span><span class="sxs-lookup"><span data-stu-id="e83e5-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="e83e5-146">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-146">Compose</span></span><br><span data-ttu-id="e83e5-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-147">Read</span></span> | [<span data-ttu-id="e83e5-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e83e5-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="e83e5-149">1,5</span><span class="sxs-lookup"><span data-stu-id="e83e5-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e83e5-150">requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="e83e5-151">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-151">Compose</span></span><br><span data-ttu-id="e83e5-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-152">Read</span></span> | [<span data-ttu-id="e83e5-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e83e5-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="e83e5-154">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e83e5-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="e83e5-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="e83e5-156">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-156">Compose</span></span><br><span data-ttu-id="e83e5-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-157">Read</span></span> | [<span data-ttu-id="e83e5-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e83e5-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="e83e5-159">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e83e5-160">ui</span><span class="sxs-lookup"><span data-stu-id="e83e5-160">ui</span></span>](#ui-ui) | <span data-ttu-id="e83e5-161">Composition</span><span class="sxs-lookup"><span data-stu-id="e83e5-161">Compose</span></span><br><span data-ttu-id="e83e5-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-162">Read</span></span> | [<span data-ttu-id="e83e5-163">UI</span><span class="sxs-lookup"><span data-stu-id="e83e5-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.8&preserve-view=true) | [<span data-ttu-id="e83e5-164">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="e83e5-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="e83e5-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="e83e5-166">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="e83e5-166">contentLanguage: String</span></span>

<span data-ttu-id="e83e5-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="e83e5-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="e83e5-168">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="e83e5-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-169">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-169">Type</span></span>

*   <span data-ttu-id="e83e5-170">String</span><span class="sxs-lookup"><span data-stu-id="e83e5-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e83e5-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-171">Requirements</span></span>

|<span data-ttu-id="e83e5-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-172">Requirement</span></span>| <span data-ttu-id="e83e5-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-175">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-175">1.1</span></span>|
|[<span data-ttu-id="e83e5-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e83e5-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="e83e5-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="e83e5-179">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="e83e5-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="e83e5-180">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e83e5-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-181">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-181">Type</span></span>

*   [<span data-ttu-id="e83e5-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e83e5-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="e83e5-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-183">Requirements</span></span>

|<span data-ttu-id="e83e5-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-184">Requirement</span></span>| <span data-ttu-id="e83e5-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-187">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-187">1.1</span></span>|
|[<span data-ttu-id="e83e5-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e83e5-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="e83e5-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="e83e5-191">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="e83e5-191">displayLanguage: String</span></span>

<span data-ttu-id="e83e5-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="e83e5-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="e83e5-193">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="e83e5-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-194">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-194">Type</span></span>

*   <span data-ttu-id="e83e5-195">String</span><span class="sxs-lookup"><span data-stu-id="e83e5-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e83e5-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-196">Requirements</span></span>

|<span data-ttu-id="e83e5-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-197">Requirement</span></span>| <span data-ttu-id="e83e5-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-200">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-200">1.1</span></span>|
|[<span data-ttu-id="e83e5-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e83e5-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="e83e5-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="e83e5-204">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="e83e5-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="e83e5-205">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="e83e5-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e83e5-206">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="e83e5-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-207">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-207">Type</span></span>

*   [<span data-ttu-id="e83e5-208">HostType</span><span class="sxs-lookup"><span data-stu-id="e83e5-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="e83e5-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-209">Requirements</span></span>

|<span data-ttu-id="e83e5-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-210">Requirement</span></span>| <span data-ttu-id="e83e5-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-213">1,5</span><span class="sxs-lookup"><span data-stu-id="e83e5-213">1.5</span></span>|
|[<span data-ttu-id="e83e5-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-215">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e83e5-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="e83e5-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="e83e5-217">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="e83e5-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="e83e5-218">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="e83e5-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="e83e5-219">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="e83e5-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-220">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-220">Type</span></span>

*   [<span data-ttu-id="e83e5-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e83e5-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="e83e5-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-222">Requirements</span></span>

|<span data-ttu-id="e83e5-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-223">Requirement</span></span>| <span data-ttu-id="e83e5-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-226">1,5</span><span class="sxs-lookup"><span data-stu-id="e83e5-226">1.5</span></span>|
|[<span data-ttu-id="e83e5-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e83e5-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="e83e5-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="e83e5-230">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="e83e5-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="e83e5-231">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="e83e5-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-232">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-232">Type</span></span>

*   [<span data-ttu-id="e83e5-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e83e5-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="e83e5-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-234">Requirements</span></span>

|<span data-ttu-id="e83e5-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-235">Requirement</span></span>| <span data-ttu-id="e83e5-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-238">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-238">1.1</span></span>|
|[<span data-ttu-id="e83e5-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-240">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e83e5-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="e83e5-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="e83e5-242">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="e83e5-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="e83e5-243">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="e83e5-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e83e5-244">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="e83e5-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-245">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-245">Type</span></span>

*   [<span data-ttu-id="e83e5-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e83e5-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e83e5-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-247">Requirements</span></span>

|<span data-ttu-id="e83e5-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-248">Requirement</span></span>| <span data-ttu-id="e83e5-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-251">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-251">1.1</span></span>|
|[<span data-ttu-id="e83e5-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="e83e5-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="e83e5-253">Restreinte</span><span class="sxs-lookup"><span data-stu-id="e83e5-253">Restricted</span></span>|
|[<span data-ttu-id="e83e5-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="e83e5-256">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="e83e5-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="e83e5-257">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="e83e5-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e83e5-258">Type</span><span class="sxs-lookup"><span data-stu-id="e83e5-258">Type</span></span>

*   [<span data-ttu-id="e83e5-259">UI</span><span class="sxs-lookup"><span data-stu-id="e83e5-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="e83e5-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="e83e5-260">Requirements</span></span>

|<span data-ttu-id="e83e5-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="e83e5-261">Requirement</span></span>| <span data-ttu-id="e83e5-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="e83e5-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="e83e5-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="e83e5-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e83e5-264">1.1</span><span class="sxs-lookup"><span data-stu-id="e83e5-264">1.1</span></span>|
|[<span data-ttu-id="e83e5-265">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="e83e5-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e83e5-266">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="e83e5-266">Compose or Read</span></span>|
