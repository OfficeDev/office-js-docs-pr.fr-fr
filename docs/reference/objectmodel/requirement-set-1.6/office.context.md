---
title: Office. Context-ensemble de conditions requises 1,6
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,6.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 55e3761aea94d902903c53a9b3be687d94b42e12
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570757"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="36dda-103">contexte (boîte aux lettres requise définie sur 1,6)</span><span class="sxs-lookup"><span data-stu-id="36dda-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="36dda-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="36dda-104">[Office](office.md).context</span></span>

<span data-ttu-id="36dda-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="36dda-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="36dda-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="36dda-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="36dda-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-107">Requirements</span></span>

|<span data-ttu-id="36dda-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-108">Requirement</span></span>| <span data-ttu-id="36dda-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-111">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-111">1.1</span></span>|
|[<span data-ttu-id="36dda-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="36dda-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="36dda-114">Properties</span></span>

| <span data-ttu-id="36dda-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="36dda-115">Property</span></span> | <span data-ttu-id="36dda-116">Modes</span><span class="sxs-lookup"><span data-stu-id="36dda-116">Modes</span></span> | <span data-ttu-id="36dda-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="36dda-117">Return type</span></span> | <span data-ttu-id="36dda-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="36dda-118">Minimum</span></span><br><span data-ttu-id="36dda-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="36dda-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="36dda-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="36dda-121">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-121">Compose</span></span><br><span data-ttu-id="36dda-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-122">Read</span></span> | <span data-ttu-id="36dda-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="36dda-123">String</span></span> | [<span data-ttu-id="36dda-124">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36dda-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="36dda-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="36dda-126">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-126">Compose</span></span><br><span data-ttu-id="36dda-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-127">Read</span></span> | [<span data-ttu-id="36dda-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="36dda-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="36dda-129">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36dda-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="36dda-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="36dda-131">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-131">Compose</span></span><br><span data-ttu-id="36dda-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-132">Read</span></span> | <span data-ttu-id="36dda-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="36dda-133">String</span></span> | [<span data-ttu-id="36dda-134">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36dda-135">hote</span><span class="sxs-lookup"><span data-stu-id="36dda-135">host</span></span>](#host-hosttype) | <span data-ttu-id="36dda-136">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-136">Compose</span></span><br><span data-ttu-id="36dda-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-137">Read</span></span> | [<span data-ttu-id="36dda-138">HostType</span><span class="sxs-lookup"><span data-stu-id="36dda-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="36dda-139">1,5</span><span class="sxs-lookup"><span data-stu-id="36dda-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="36dda-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="36dda-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="36dda-141">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-141">Compose</span></span><br><span data-ttu-id="36dda-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-142">Read</span></span> | [<span data-ttu-id="36dda-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="36dda-144">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36dda-145">plateforme</span><span class="sxs-lookup"><span data-stu-id="36dda-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="36dda-146">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-146">Compose</span></span><br><span data-ttu-id="36dda-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-147">Read</span></span> | [<span data-ttu-id="36dda-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="36dda-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="36dda-149">1,5</span><span class="sxs-lookup"><span data-stu-id="36dda-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="36dda-150">requise</span><span class="sxs-lookup"><span data-stu-id="36dda-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="36dda-151">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-151">Compose</span></span><br><span data-ttu-id="36dda-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-152">Read</span></span> | [<span data-ttu-id="36dda-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="36dda-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="36dda-154">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36dda-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="36dda-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="36dda-156">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-156">Compose</span></span><br><span data-ttu-id="36dda-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-157">Read</span></span> | [<span data-ttu-id="36dda-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="36dda-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="36dda-159">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36dda-160">ui</span><span class="sxs-lookup"><span data-stu-id="36dda-160">ui</span></span>](#ui-ui) | <span data-ttu-id="36dda-161">Composition</span><span class="sxs-lookup"><span data-stu-id="36dda-161">Compose</span></span><br><span data-ttu-id="36dda-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-162">Read</span></span> | [<span data-ttu-id="36dda-163">UI</span><span class="sxs-lookup"><span data-stu-id="36dda-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="36dda-164">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="36dda-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="36dda-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="36dda-166">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="36dda-166">contentLanguage: String</span></span>

<span data-ttu-id="36dda-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="36dda-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="36dda-168">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="36dda-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-169">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-169">Type</span></span>

*   <span data-ttu-id="36dda-170">String</span><span class="sxs-lookup"><span data-stu-id="36dda-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36dda-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-171">Requirements</span></span>

|<span data-ttu-id="36dda-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-172">Requirement</span></span>| <span data-ttu-id="36dda-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-175">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-175">1.1</span></span>|
|[<span data-ttu-id="36dda-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36dda-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="36dda-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="36dda-179">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="36dda-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="36dda-180">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="36dda-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-181">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-181">Type</span></span>

*   [<span data-ttu-id="36dda-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="36dda-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="36dda-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-183">Requirements</span></span>

|<span data-ttu-id="36dda-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-184">Requirement</span></span>| <span data-ttu-id="36dda-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-187">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-187">1.1</span></span>|
|[<span data-ttu-id="36dda-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36dda-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="36dda-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="36dda-191">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="36dda-191">displayLanguage: String</span></span>

<span data-ttu-id="36dda-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="36dda-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="36dda-193">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="36dda-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-194">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-194">Type</span></span>

*   <span data-ttu-id="36dda-195">String</span><span class="sxs-lookup"><span data-stu-id="36dda-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36dda-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-196">Requirements</span></span>

|<span data-ttu-id="36dda-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-197">Requirement</span></span>| <span data-ttu-id="36dda-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-200">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-200">1.1</span></span>|
|[<span data-ttu-id="36dda-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36dda-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="36dda-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="36dda-204">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="36dda-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="36dda-205">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="36dda-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="36dda-206">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="36dda-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-207">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-207">Type</span></span>

*   [<span data-ttu-id="36dda-208">HostType</span><span class="sxs-lookup"><span data-stu-id="36dda-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="36dda-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-209">Requirements</span></span>

|<span data-ttu-id="36dda-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-210">Requirement</span></span>| <span data-ttu-id="36dda-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-213">1,5</span><span class="sxs-lookup"><span data-stu-id="36dda-213">1.5</span></span>|
|[<span data-ttu-id="36dda-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-215">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36dda-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="36dda-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="36dda-217">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="36dda-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="36dda-218">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="36dda-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="36dda-219">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="36dda-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-220">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-220">Type</span></span>

*   [<span data-ttu-id="36dda-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="36dda-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="36dda-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-222">Requirements</span></span>

|<span data-ttu-id="36dda-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-223">Requirement</span></span>| <span data-ttu-id="36dda-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-226">1,5</span><span class="sxs-lookup"><span data-stu-id="36dda-226">1.5</span></span>|
|[<span data-ttu-id="36dda-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36dda-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="36dda-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="36dda-230">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="36dda-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="36dda-231">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="36dda-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-232">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-232">Type</span></span>

*   [<span data-ttu-id="36dda-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="36dda-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="36dda-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-234">Requirements</span></span>

|<span data-ttu-id="36dda-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-235">Requirement</span></span>| <span data-ttu-id="36dda-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-238">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-238">1.1</span></span>|
|[<span data-ttu-id="36dda-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-240">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36dda-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="36dda-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="36dda-242">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="36dda-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="36dda-243">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="36dda-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="36dda-244">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="36dda-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-245">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-245">Type</span></span>

*   [<span data-ttu-id="36dda-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="36dda-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="36dda-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-247">Requirements</span></span>

|<span data-ttu-id="36dda-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-248">Requirement</span></span>| <span data-ttu-id="36dda-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-251">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-251">1.1</span></span>|
|[<span data-ttu-id="36dda-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="36dda-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="36dda-253">Restreinte</span><span class="sxs-lookup"><span data-stu-id="36dda-253">Restricted</span></span>|
|[<span data-ttu-id="36dda-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="36dda-256">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="36dda-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="36dda-257">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="36dda-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="36dda-258">Type</span><span class="sxs-lookup"><span data-stu-id="36dda-258">Type</span></span>

*   [<span data-ttu-id="36dda-259">UI</span><span class="sxs-lookup"><span data-stu-id="36dda-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="36dda-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="36dda-260">Requirements</span></span>

|<span data-ttu-id="36dda-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="36dda-261">Requirement</span></span>| <span data-ttu-id="36dda-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="36dda-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="36dda-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="36dda-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36dda-264">1.1</span><span class="sxs-lookup"><span data-stu-id="36dda-264">1.1</span></span>|
|[<span data-ttu-id="36dda-265">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="36dda-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36dda-266">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="36dda-266">Compose or Read</span></span>|
