---
title: Office.context - ensemble de conditions requises 1.5
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.5.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 8aedd711665d902cf3cc733901df9e3a3cc86886
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591008"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="6c7ce-103">context (ensemble de conditions requises de boîte aux lettres 1.5)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="6c7ce-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="6c7ce-104">[Office](office.md).context</span></span>

<span data-ttu-id="6c7ce-105">Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="6c7ce-106">Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir la liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c7ce-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-107">Requirements</span></span>

|<span data-ttu-id="6c7ce-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-108">Requirement</span></span>| <span data-ttu-id="6c7ce-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-111">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-111">1.1</span></span>|
|[<span data-ttu-id="6c7ce-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="6c7ce-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="6c7ce-114">Properties</span></span>

| <span data-ttu-id="6c7ce-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="6c7ce-115">Property</span></span> | <span data-ttu-id="6c7ce-116">Modes</span><span class="sxs-lookup"><span data-stu-id="6c7ce-116">Modes</span></span> | <span data-ttu-id="6c7ce-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="6c7ce-117">Return type</span></span> | <span data-ttu-id="6c7ce-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="6c7ce-118">Minimum</span></span><br><span data-ttu-id="6c7ce-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6c7ce-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="6c7ce-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="6c7ce-121">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-121">Compose</span></span><br><span data-ttu-id="6c7ce-122">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-122">Read</span></span> | <span data-ttu-id="6c7ce-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6c7ce-123">String</span></span> | [<span data-ttu-id="6c7ce-124">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6c7ce-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="6c7ce-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="6c7ce-126">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-126">Compose</span></span><br><span data-ttu-id="6c7ce-127">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-127">Read</span></span> | [<span data-ttu-id="6c7ce-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6c7ce-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6c7ce-129">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6c7ce-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="6c7ce-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="6c7ce-131">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-131">Compose</span></span><br><span data-ttu-id="6c7ce-132">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-132">Read</span></span> | <span data-ttu-id="6c7ce-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6c7ce-133">String</span></span> | [<span data-ttu-id="6c7ce-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6c7ce-135">host</span><span class="sxs-lookup"><span data-stu-id="6c7ce-135">host</span></span>](#host-hosttype) | <span data-ttu-id="6c7ce-136">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-136">Compose</span></span><br><span data-ttu-id="6c7ce-137">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-137">Read</span></span> | [<span data-ttu-id="6c7ce-138">HostType</span><span class="sxs-lookup"><span data-stu-id="6c7ce-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6c7ce-139">1.5</span><span class="sxs-lookup"><span data-stu-id="6c7ce-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6c7ce-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="6c7ce-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="6c7ce-141">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-141">Compose</span></span><br><span data-ttu-id="6c7ce-142">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-142">Read</span></span> | [<span data-ttu-id="6c7ce-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6c7ce-144">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6c7ce-145">platform</span><span class="sxs-lookup"><span data-stu-id="6c7ce-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="6c7ce-146">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-146">Compose</span></span><br><span data-ttu-id="6c7ce-147">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-147">Read</span></span> | [<span data-ttu-id="6c7ce-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6c7ce-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6c7ce-149">1.5</span><span class="sxs-lookup"><span data-stu-id="6c7ce-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6c7ce-150">requirements</span><span class="sxs-lookup"><span data-stu-id="6c7ce-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="6c7ce-151">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-151">Compose</span></span><br><span data-ttu-id="6c7ce-152">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-152">Read</span></span> | [<span data-ttu-id="6c7ce-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6c7ce-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6c7ce-154">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6c7ce-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="6c7ce-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="6c7ce-156">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-156">Compose</span></span><br><span data-ttu-id="6c7ce-157">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-157">Read</span></span> | [<span data-ttu-id="6c7ce-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6c7ce-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6c7ce-159">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6c7ce-160">ui</span><span class="sxs-lookup"><span data-stu-id="6c7ce-160">ui</span></span>](#ui-ui) | <span data-ttu-id="6c7ce-161">Composition</span><span class="sxs-lookup"><span data-stu-id="6c7ce-161">Compose</span></span><br><span data-ttu-id="6c7ce-162">Lire</span><span class="sxs-lookup"><span data-stu-id="6c7ce-162">Read</span></span> | [<span data-ttu-id="6c7ce-163">UI</span><span class="sxs-lookup"><span data-stu-id="6c7ce-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="6c7ce-164">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="6c7ce-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="6c7ce-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="6c7ce-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6c7ce-166">contentLanguage: String</span></span>

<span data-ttu-id="6c7ce-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="6c7ce-168">La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options de > langue dans l Office `contentLanguage` application cliente.  </span><span class="sxs-lookup"><span data-stu-id="6c7ce-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-169">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-169">Type</span></span>

*   <span data-ttu-id="6c7ce-170">String</span><span class="sxs-lookup"><span data-stu-id="6c7ce-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c7ce-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-171">Requirements</span></span>

|<span data-ttu-id="6c7ce-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-172">Requirement</span></span>| <span data-ttu-id="6c7ce-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-175">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-175">1.1</span></span>|
|[<span data-ttu-id="6c7ce-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c7ce-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="6c7ce-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="6c7ce-179">diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="6c7ce-180">Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-181">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-181">Type</span></span>

*   [<span data-ttu-id="6c7ce-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6c7ce-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="6c7ce-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-183">Requirements</span></span>

|<span data-ttu-id="6c7ce-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-184">Requirement</span></span>| <span data-ttu-id="6c7ce-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-187">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-187">1.1</span></span>|
|[<span data-ttu-id="6c7ce-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c7ce-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="6c7ce-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="6c7ce-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6c7ce-191">displayLanguage: String</span></span>

<span data-ttu-id="6c7ce-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="6c7ce-193">La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office..  </span><span class="sxs-lookup"><span data-stu-id="6c7ce-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-194">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-194">Type</span></span>

*   <span data-ttu-id="6c7ce-195">String</span><span class="sxs-lookup"><span data-stu-id="6c7ce-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6c7ce-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-196">Requirements</span></span>

|<span data-ttu-id="6c7ce-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-197">Requirement</span></span>| <span data-ttu-id="6c7ce-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-200">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-200">1.1</span></span>|
|[<span data-ttu-id="6c7ce-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c7ce-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="6c7ce-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="6c7ce-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="6c7ce-205">Obtient Office application qui héberge le module.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="6c7ce-206">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-207">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-207">Type</span></span>

*   [<span data-ttu-id="6c7ce-208">HostType</span><span class="sxs-lookup"><span data-stu-id="6c7ce-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="6c7ce-209">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-209">Requirements</span></span>

|<span data-ttu-id="6c7ce-210">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-210">Requirement</span></span>| <span data-ttu-id="6c7ce-211">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-212">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-213">1,5</span><span class="sxs-lookup"><span data-stu-id="6c7ce-213">1.5</span></span>|
|[<span data-ttu-id="6c7ce-214">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-215">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c7ce-216">Exemple</span><span class="sxs-lookup"><span data-stu-id="6c7ce-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="6c7ce-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="6c7ce-218">Fournit la plateforme sur laquelle le module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="6c7ce-219">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-220">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-220">Type</span></span>

*   [<span data-ttu-id="6c7ce-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6c7ce-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="6c7ce-222">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-222">Requirements</span></span>

|<span data-ttu-id="6c7ce-223">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-223">Requirement</span></span>| <span data-ttu-id="6c7ce-224">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-225">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-226">1,5</span><span class="sxs-lookup"><span data-stu-id="6c7ce-226">1.5</span></span>|
|[<span data-ttu-id="6c7ce-227">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-228">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c7ce-229">Exemple</span><span class="sxs-lookup"><span data-stu-id="6c7ce-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="6c7ce-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="6c7ce-231">Fournit une méthode pour déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-232">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-232">Type</span></span>

*   [<span data-ttu-id="6c7ce-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6c7ce-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="6c7ce-234">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-234">Requirements</span></span>

|<span data-ttu-id="6c7ce-235">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-235">Requirement</span></span>| <span data-ttu-id="6c7ce-236">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-237">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-238">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-238">1.1</span></span>|
|[<span data-ttu-id="6c7ce-239">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-240">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6c7ce-241">Exemple</span><span class="sxs-lookup"><span data-stu-id="6c7ce-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="6c7ce-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="6c7ce-243">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6c7ce-244">L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-245">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-245">Type</span></span>

*   [<span data-ttu-id="6c7ce-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6c7ce-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6c7ce-247">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-247">Requirements</span></span>

|<span data-ttu-id="6c7ce-248">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-248">Requirement</span></span>| <span data-ttu-id="6c7ce-249">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-250">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-251">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-251">1.1</span></span>|
|[<span data-ttu-id="6c7ce-252">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6c7ce-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="6c7ce-253">Restreinte</span><span class="sxs-lookup"><span data-stu-id="6c7ce-253">Restricted</span></span>|
|[<span data-ttu-id="6c7ce-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="6c7ce-256">Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="6c7ce-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="6c7ce-257">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.</span><span class="sxs-lookup"><span data-stu-id="6c7ce-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6c7ce-258">Type</span><span class="sxs-lookup"><span data-stu-id="6c7ce-258">Type</span></span>

*   [<span data-ttu-id="6c7ce-259">UI</span><span class="sxs-lookup"><span data-stu-id="6c7ce-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="6c7ce-260">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6c7ce-260">Requirements</span></span>

|<span data-ttu-id="6c7ce-261">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6c7ce-261">Requirement</span></span>| <span data-ttu-id="6c7ce-262">Valeur</span><span class="sxs-lookup"><span data-stu-id="6c7ce-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="6c7ce-263">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6c7ce-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6c7ce-264">1.1</span><span class="sxs-lookup"><span data-stu-id="6c7ce-264">1.1</span></span>|
|[<span data-ttu-id="6c7ce-265">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6c7ce-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6c7ce-266">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="6c7ce-266">Compose or Read</span></span>|
