---
title: Office. Context-ensemble de conditions requises 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 81f2fa0af17e29b1be8274451027b20a8ccb4c3c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814583"
---
# <a name="context"></a><span data-ttu-id="ffdaa-102">context</span><span class="sxs-lookup"><span data-stu-id="ffdaa-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="ffdaa-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ffdaa-103">[Office](office.md).context</span></span>

<span data-ttu-id="ffdaa-104">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ffdaa-105">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.7).</span><span class="sxs-lookup"><span data-stu-id="ffdaa-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.7).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ffdaa-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-106">Requirements</span></span>

|<span data-ttu-id="ffdaa-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-107">Requirement</span></span>| <span data-ttu-id="ffdaa-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-110">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-110">1.1</span></span>|
|[<span data-ttu-id="ffdaa-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ffdaa-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ffdaa-113">Properties</span></span>

| <span data-ttu-id="ffdaa-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="ffdaa-114">Property</span></span> | <span data-ttu-id="ffdaa-115">Modes</span><span class="sxs-lookup"><span data-stu-id="ffdaa-115">Modes</span></span> | <span data-ttu-id="ffdaa-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ffdaa-116">Return type</span></span> | <span data-ttu-id="ffdaa-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="ffdaa-117">Minimum</span></span><br><span data-ttu-id="ffdaa-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ffdaa-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ffdaa-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ffdaa-120">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-120">Compose</span></span><br><span data-ttu-id="ffdaa-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-121">Read</span></span> | <span data-ttu-id="ffdaa-122">String</span><span class="sxs-lookup"><span data-stu-id="ffdaa-122">String</span></span> | [<span data-ttu-id="ffdaa-123">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-124">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="ffdaa-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ffdaa-125">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-125">Compose</span></span><br><span data-ttu-id="ffdaa-126">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-126">Read</span></span> | [<span data-ttu-id="ffdaa-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ffdaa-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.7) | [<span data-ttu-id="ffdaa-128">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ffdaa-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ffdaa-130">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-130">Compose</span></span><br><span data-ttu-id="ffdaa-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-131">Read</span></span> | <span data-ttu-id="ffdaa-132">String</span><span class="sxs-lookup"><span data-stu-id="ffdaa-132">String</span></span> | [<span data-ttu-id="ffdaa-133">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-134">hote</span><span class="sxs-lookup"><span data-stu-id="ffdaa-134">host</span></span>](#host-hosttype) | <span data-ttu-id="ffdaa-135">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-135">Compose</span></span><br><span data-ttu-id="ffdaa-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-136">Read</span></span> | [<span data-ttu-id="ffdaa-137">HostType</span><span class="sxs-lookup"><span data-stu-id="ffdaa-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.7) | [<span data-ttu-id="ffdaa-138">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="ffdaa-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ffdaa-140">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-140">Compose</span></span><br><span data-ttu-id="ffdaa-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-141">Read</span></span> | [<span data-ttu-id="ffdaa-142">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-142">Mailbox</span></span>](/javascript/api/office/office.mailbox?view=outlook-js-1.7) | [<span data-ttu-id="ffdaa-143">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-144">plateforme</span><span class="sxs-lookup"><span data-stu-id="ffdaa-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ffdaa-145">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-145">Compose</span></span><br><span data-ttu-id="ffdaa-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-146">Read</span></span> | [<span data-ttu-id="ffdaa-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ffdaa-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.7) | [<span data-ttu-id="ffdaa-148">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-149">requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ffdaa-150">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-150">Compose</span></span><br><span data-ttu-id="ffdaa-151">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-151">Read</span></span> | [<span data-ttu-id="ffdaa-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ffdaa-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.7) | [<span data-ttu-id="ffdaa-153">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ffdaa-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ffdaa-155">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-155">Compose</span></span><br><span data-ttu-id="ffdaa-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-156">Read</span></span> | [<span data-ttu-id="ffdaa-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ffdaa-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.7) | [<span data-ttu-id="ffdaa-158">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ffdaa-159">ui</span><span class="sxs-lookup"><span data-stu-id="ffdaa-159">ui</span></span>](#ui-ui) | <span data-ttu-id="ffdaa-160">Composition</span><span class="sxs-lookup"><span data-stu-id="ffdaa-160">Compose</span></span><br><span data-ttu-id="ffdaa-161">Lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-161">Read</span></span> | [<span data-ttu-id="ffdaa-162">UI</span><span class="sxs-lookup"><span data-stu-id="ffdaa-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.7) | [<span data-ttu-id="ffdaa-163">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ffdaa-164">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="ffdaa-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="ffdaa-165">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="ffdaa-165">contentLanguage: String</span></span>

<span data-ttu-id="ffdaa-166">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ffdaa-167">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-168">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-168">Type</span></span>

*   <span data-ttu-id="ffdaa-169">String</span><span class="sxs-lookup"><span data-stu-id="ffdaa-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ffdaa-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-170">Requirements</span></span>

|<span data-ttu-id="ffdaa-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-171">Requirement</span></span>| <span data-ttu-id="ffdaa-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-174">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-174">1.1</span></span>|
|[<span data-ttu-id="ffdaa-175">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-175">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-176">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ffdaa-177">Exemple</span><span class="sxs-lookup"><span data-stu-id="ffdaa-177">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="ffdaa-178">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ffdaa-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ffdaa-179">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-180">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-180">Type</span></span>

*   [<span data-ttu-id="ffdaa-181">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ffdaa-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ffdaa-182">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-182">Requirements</span></span>

|<span data-ttu-id="ffdaa-183">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-183">Requirement</span></span>| <span data-ttu-id="ffdaa-184">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-185">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-186">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-186">1.1</span></span>|
|[<span data-ttu-id="ffdaa-187">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-187">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-188">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ffdaa-189">Exemple</span><span class="sxs-lookup"><span data-stu-id="ffdaa-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ffdaa-190">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="ffdaa-190">displayLanguage: String</span></span>

<span data-ttu-id="ffdaa-191">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="ffdaa-192">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-193">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-193">Type</span></span>

*   <span data-ttu-id="ffdaa-194">String</span><span class="sxs-lookup"><span data-stu-id="ffdaa-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ffdaa-195">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-195">Requirements</span></span>

|<span data-ttu-id="ffdaa-196">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-196">Requirement</span></span>| <span data-ttu-id="ffdaa-197">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-198">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-199">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-199">1.1</span></span>|
|[<span data-ttu-id="ffdaa-200">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-200">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-201">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ffdaa-202">Exemple</span><span class="sxs-lookup"><span data-stu-id="ffdaa-202">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="ffdaa-203">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ffdaa-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ffdaa-204">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-205">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-205">Type</span></span>

*   [<span data-ttu-id="ffdaa-206">HostType</span><span class="sxs-lookup"><span data-stu-id="ffdaa-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ffdaa-207">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-207">Requirements</span></span>

|<span data-ttu-id="ffdaa-208">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-208">Requirement</span></span>| <span data-ttu-id="ffdaa-209">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-210">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-211">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-211">1.1</span></span>|
|[<span data-ttu-id="ffdaa-212">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-212">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-213">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ffdaa-214">Exemple</span><span class="sxs-lookup"><span data-stu-id="ffdaa-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="ffdaa-215">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ffdaa-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ffdaa-216">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-217">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-217">Type</span></span>

*   [<span data-ttu-id="ffdaa-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ffdaa-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ffdaa-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-219">Requirements</span></span>

|<span data-ttu-id="ffdaa-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-220">Requirement</span></span>| <span data-ttu-id="ffdaa-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-223">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-223">1.1</span></span>|
|[<span data-ttu-id="ffdaa-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-224">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ffdaa-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="ffdaa-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="ffdaa-227">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ffdaa-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ffdaa-228">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-229">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-229">Type</span></span>

*   [<span data-ttu-id="ffdaa-230">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ffdaa-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ffdaa-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-231">Requirements</span></span>

|<span data-ttu-id="ffdaa-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-232">Requirement</span></span>| <span data-ttu-id="ffdaa-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-235">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-235">1.1</span></span>|
|[<span data-ttu-id="ffdaa-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-236">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ffdaa-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="ffdaa-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="ffdaa-239">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ffdaa-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ffdaa-240">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ffdaa-241">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-242">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-242">Type</span></span>

*   [<span data-ttu-id="ffdaa-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ffdaa-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ffdaa-244">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-244">Requirements</span></span>

|<span data-ttu-id="ffdaa-245">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-245">Requirement</span></span>| <span data-ttu-id="ffdaa-246">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-247">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-248">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-248">1.1</span></span>|
|[<span data-ttu-id="ffdaa-249">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ffdaa-249">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ffdaa-250">Restreinte</span><span class="sxs-lookup"><span data-stu-id="ffdaa-250">Restricted</span></span>|
|[<span data-ttu-id="ffdaa-251">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-251">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-252">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="ffdaa-253">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ffdaa-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ffdaa-254">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="ffdaa-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ffdaa-255">Type</span><span class="sxs-lookup"><span data-stu-id="ffdaa-255">Type</span></span>

*   [<span data-ttu-id="ffdaa-256">UI</span><span class="sxs-lookup"><span data-stu-id="ffdaa-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ffdaa-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ffdaa-257">Requirements</span></span>

|<span data-ttu-id="ffdaa-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ffdaa-258">Requirement</span></span>| <span data-ttu-id="ffdaa-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="ffdaa-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="ffdaa-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ffdaa-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ffdaa-261">1.1</span><span class="sxs-lookup"><span data-stu-id="ffdaa-261">1.1</span></span>|
|[<span data-ttu-id="ffdaa-262">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ffdaa-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ffdaa-263">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ffdaa-263">Compose or Read</span></span>|
