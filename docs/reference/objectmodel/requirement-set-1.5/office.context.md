---
title: Office. Context-ensemble de conditions requises 1,5
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 17bbe15bc1fd697756caba996b2d4bfbbac15d2e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44608655"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="1ebe9-103">contexte (boîte aux lettres requise définie sur 1,5)</span><span class="sxs-lookup"><span data-stu-id="1ebe9-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="1ebe9-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="1ebe9-104">[Office](office.md).context</span></span>

<span data-ttu-id="1ebe9-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="1ebe9-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.5).</span><span class="sxs-lookup"><span data-stu-id="1ebe9-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ebe9-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-107">Requirements</span></span>

|<span data-ttu-id="1ebe9-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-108">Requirement</span></span>| <span data-ttu-id="1ebe9-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-111">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-111">1.1</span></span>|
|[<span data-ttu-id="1ebe9-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1ebe9-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="1ebe9-114">Properties</span></span>

| <span data-ttu-id="1ebe9-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="1ebe9-115">Property</span></span> | <span data-ttu-id="1ebe9-116">Modes</span><span class="sxs-lookup"><span data-stu-id="1ebe9-116">Modes</span></span> | <span data-ttu-id="1ebe9-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="1ebe9-117">Return type</span></span> | <span data-ttu-id="1ebe9-118">Minimale</span><span class="sxs-lookup"><span data-stu-id="1ebe9-118">Minimum</span></span><br><span data-ttu-id="1ebe9-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1ebe9-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="1ebe9-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="1ebe9-121">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-121">Compose</span></span><br><span data-ttu-id="1ebe9-122">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-122">Read</span></span> | <span data-ttu-id="1ebe9-123">String</span><span class="sxs-lookup"><span data-stu-id="1ebe9-123">String</span></span> | [<span data-ttu-id="1ebe9-124">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="1ebe9-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="1ebe9-126">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-126">Compose</span></span><br><span data-ttu-id="1ebe9-127">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-127">Read</span></span> | [<span data-ttu-id="1ebe9-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1ebe9-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5) | [<span data-ttu-id="1ebe9-129">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="1ebe9-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="1ebe9-131">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-131">Compose</span></span><br><span data-ttu-id="1ebe9-132">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-132">Read</span></span> | <span data-ttu-id="1ebe9-133">String</span><span class="sxs-lookup"><span data-stu-id="1ebe9-133">String</span></span> | [<span data-ttu-id="1ebe9-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-135">hote</span><span class="sxs-lookup"><span data-stu-id="1ebe9-135">host</span></span>](#host-hosttype) | <span data-ttu-id="1ebe9-136">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-136">Compose</span></span><br><span data-ttu-id="1ebe9-137">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-137">Read</span></span> | [<span data-ttu-id="1ebe9-138">HostType</span><span class="sxs-lookup"><span data-stu-id="1ebe9-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5) | [<span data-ttu-id="1ebe9-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="1ebe9-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="1ebe9-141">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-141">Compose</span></span><br><span data-ttu-id="1ebe9-142">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-142">Read</span></span> | [<span data-ttu-id="1ebe9-143">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5) | [<span data-ttu-id="1ebe9-144">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-145">plateforme</span><span class="sxs-lookup"><span data-stu-id="1ebe9-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="1ebe9-146">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-146">Compose</span></span><br><span data-ttu-id="1ebe9-147">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-147">Read</span></span> | [<span data-ttu-id="1ebe9-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1ebe9-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5) | [<span data-ttu-id="1ebe9-149">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-150">requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="1ebe9-151">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-151">Compose</span></span><br><span data-ttu-id="1ebe9-152">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-152">Read</span></span> | [<span data-ttu-id="1ebe9-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1ebe9-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5) | [<span data-ttu-id="1ebe9-154">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="1ebe9-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="1ebe9-156">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-156">Compose</span></span><br><span data-ttu-id="1ebe9-157">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-157">Read</span></span> | [<span data-ttu-id="1ebe9-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1ebe9-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5) | [<span data-ttu-id="1ebe9-159">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1ebe9-160">ui</span><span class="sxs-lookup"><span data-stu-id="1ebe9-160">ui</span></span>](#ui-ui) | <span data-ttu-id="1ebe9-161">Composition</span><span class="sxs-lookup"><span data-stu-id="1ebe9-161">Compose</span></span><br><span data-ttu-id="1ebe9-162">Read</span><span class="sxs-lookup"><span data-stu-id="1ebe9-162">Read</span></span> | [<span data-ttu-id="1ebe9-163">UI</span><span class="sxs-lookup"><span data-stu-id="1ebe9-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5) | [<span data-ttu-id="1ebe9-164">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="1ebe9-165">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="1ebe9-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="1ebe9-166">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="1ebe9-166">contentLanguage: String</span></span>

<span data-ttu-id="1ebe9-167">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="1ebe9-168">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-169">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-169">Type</span></span>

*   <span data-ttu-id="1ebe9-170">String</span><span class="sxs-lookup"><span data-stu-id="1ebe9-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ebe9-171">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-171">Requirements</span></span>

|<span data-ttu-id="1ebe9-172">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-172">Requirement</span></span>| <span data-ttu-id="1ebe9-173">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-174">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-175">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-175">1.1</span></span>|
|[<span data-ttu-id="1ebe9-176">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-177">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ebe9-178">Exemple</span><span class="sxs-lookup"><span data-stu-id="1ebe9-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="1ebe9-179">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1ebe9-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="1ebe9-180">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-181">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-181">Type</span></span>

*   [<span data-ttu-id="1ebe9-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1ebe9-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="1ebe9-183">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-183">Requirements</span></span>

|<span data-ttu-id="1ebe9-184">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-184">Requirement</span></span>| <span data-ttu-id="1ebe9-185">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-186">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-187">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-187">1.1</span></span>|
|[<span data-ttu-id="1ebe9-188">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-189">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ebe9-190">Exemple</span><span class="sxs-lookup"><span data-stu-id="1ebe9-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="1ebe9-191">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="1ebe9-191">displayLanguage: String</span></span>

<span data-ttu-id="1ebe9-192">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="1ebe9-193">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-194">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-194">Type</span></span>

*   <span data-ttu-id="1ebe9-195">String</span><span class="sxs-lookup"><span data-stu-id="1ebe9-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1ebe9-196">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-196">Requirements</span></span>

|<span data-ttu-id="1ebe9-197">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-197">Requirement</span></span>| <span data-ttu-id="1ebe9-198">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-199">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-200">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-200">1.1</span></span>|
|[<span data-ttu-id="1ebe9-201">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-202">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ebe9-203">Exemple</span><span class="sxs-lookup"><span data-stu-id="1ebe9-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="1ebe9-204">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="1ebe9-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="1ebe9-205">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-206">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-206">Type</span></span>

*   [<span data-ttu-id="1ebe9-207">HostType</span><span class="sxs-lookup"><span data-stu-id="1ebe9-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="1ebe9-208">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-208">Requirements</span></span>

|<span data-ttu-id="1ebe9-209">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-209">Requirement</span></span>| <span data-ttu-id="1ebe9-210">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-211">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-212">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-212">1.1</span></span>|
|[<span data-ttu-id="1ebe9-213">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-214">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ebe9-215">Exemple</span><span class="sxs-lookup"><span data-stu-id="1ebe9-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="1ebe9-216">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="1ebe9-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="1ebe9-217">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-218">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-218">Type</span></span>

*   [<span data-ttu-id="1ebe9-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1ebe9-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="1ebe9-220">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-220">Requirements</span></span>

|<span data-ttu-id="1ebe9-221">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-221">Requirement</span></span>| <span data-ttu-id="1ebe9-222">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-223">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-224">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-224">1.1</span></span>|
|[<span data-ttu-id="1ebe9-225">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-226">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ebe9-227">Exemple</span><span class="sxs-lookup"><span data-stu-id="1ebe9-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="1ebe9-228">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="1ebe9-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="1ebe9-229">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-230">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-230">Type</span></span>

*   [<span data-ttu-id="1ebe9-231">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1ebe9-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="1ebe9-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-232">Requirements</span></span>

|<span data-ttu-id="1ebe9-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-233">Requirement</span></span>| <span data-ttu-id="1ebe9-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-236">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-236">1.1</span></span>|
|[<span data-ttu-id="1ebe9-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1ebe9-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="1ebe9-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="1ebe9-240">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="1ebe9-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="1ebe9-241">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1ebe9-242">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-243">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-243">Type</span></span>

*   [<span data-ttu-id="1ebe9-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1ebe9-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1ebe9-245">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-245">Requirements</span></span>

|<span data-ttu-id="1ebe9-246">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-246">Requirement</span></span>| <span data-ttu-id="1ebe9-247">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-248">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-249">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-249">1.1</span></span>|
|[<span data-ttu-id="1ebe9-250">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="1ebe9-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="1ebe9-251">Restreinte</span><span class="sxs-lookup"><span data-stu-id="1ebe9-251">Restricted</span></span>|
|[<span data-ttu-id="1ebe9-252">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-253">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="1ebe9-254">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="1ebe9-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="1ebe9-255">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="1ebe9-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1ebe9-256">Type</span><span class="sxs-lookup"><span data-stu-id="1ebe9-256">Type</span></span>

*   [<span data-ttu-id="1ebe9-257">UI</span><span class="sxs-lookup"><span data-stu-id="1ebe9-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="1ebe9-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="1ebe9-258">Requirements</span></span>

|<span data-ttu-id="1ebe9-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="1ebe9-259">Requirement</span></span>| <span data-ttu-id="1ebe9-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="1ebe9-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="1ebe9-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="1ebe9-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1ebe9-262">1.1</span><span class="sxs-lookup"><span data-stu-id="1ebe9-262">1.1</span></span>|
|[<span data-ttu-id="1ebe9-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="1ebe9-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1ebe9-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="1ebe9-264">Compose or Read</span></span>|
