---
title: Office.context-ensemble de conditions requises 1.2
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,2.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 1b697cbe29be7d0af6fec65e47d080ebd1af17ae
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570778"
---
# <a name="context-mailbox-requirement-set-12"></a><span data-ttu-id="ae216-103">contexte (boîte aux lettres requise définie sur 1,2)</span><span class="sxs-lookup"><span data-stu-id="ae216-103">context (Mailbox requirement set 1.2)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ae216-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ae216-104">[Office](office.md).context</span></span>

<span data-ttu-id="ae216-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="ae216-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ae216-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="ae216-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae216-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ae216-107">Requirements</span></span>

|<span data-ttu-id="ae216-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-108">Requirement</span></span>| <span data-ttu-id="ae216-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="ae216-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae216-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae216-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-111">1.1</span></span>|
|[<span data-ttu-id="ae216-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ae216-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae216-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ae216-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ae216-114">Properties</span></span>

| <span data-ttu-id="ae216-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="ae216-115">Property</span></span> | <span data-ttu-id="ae216-116">Modes</span><span class="sxs-lookup"><span data-stu-id="ae216-116">Modes</span></span> | <span data-ttu-id="ae216-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ae216-117">Return type</span></span> | <span data-ttu-id="ae216-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="ae216-118">Minimum</span></span><br><span data-ttu-id="ae216-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ae216-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ae216-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ae216-121">Composition</span><span class="sxs-lookup"><span data-stu-id="ae216-121">Compose</span></span><br><span data-ttu-id="ae216-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-122">Read</span></span> | <span data-ttu-id="ae216-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ae216-123">String</span></span> | [<span data-ttu-id="ae216-124">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae216-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="ae216-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ae216-126">Composition</span><span class="sxs-lookup"><span data-stu-id="ae216-126">Compose</span></span><br><span data-ttu-id="ae216-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-127">Read</span></span> | [<span data-ttu-id="ae216-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ae216-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="ae216-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae216-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ae216-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ae216-131">Composition</span><span class="sxs-lookup"><span data-stu-id="ae216-131">Compose</span></span><br><span data-ttu-id="ae216-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-132">Read</span></span> | <span data-ttu-id="ae216-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="ae216-133">String</span></span> | [<span data-ttu-id="ae216-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae216-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="ae216-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ae216-136">Composition</span><span class="sxs-lookup"><span data-stu-id="ae216-136">Compose</span></span><br><span data-ttu-id="ae216-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-137">Read</span></span> | [<span data-ttu-id="ae216-138">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="ae216-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae216-140">requise</span><span class="sxs-lookup"><span data-stu-id="ae216-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ae216-141">Composition</span><span class="sxs-lookup"><span data-stu-id="ae216-141">Compose</span></span><br><span data-ttu-id="ae216-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-142">Read</span></span> | [<span data-ttu-id="ae216-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ae216-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="ae216-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae216-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae216-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ae216-146">Composition</span><span class="sxs-lookup"><span data-stu-id="ae216-146">Compose</span></span><br><span data-ttu-id="ae216-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-147">Read</span></span> | [<span data-ttu-id="ae216-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae216-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="ae216-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ae216-150">ui</span><span class="sxs-lookup"><span data-stu-id="ae216-150">ui</span></span>](#ui-ui) | <span data-ttu-id="ae216-151">Composition</span><span class="sxs-lookup"><span data-stu-id="ae216-151">Compose</span></span><br><span data-ttu-id="ae216-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-152">Read</span></span> | [<span data-ttu-id="ae216-153">UI</span><span class="sxs-lookup"><span data-stu-id="ae216-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2&preserve-view=true) | [<span data-ttu-id="ae216-154">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ae216-155">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="ae216-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="ae216-156">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="ae216-156">contentLanguage: String</span></span>

<span data-ttu-id="ae216-157">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="ae216-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ae216-158">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="ae216-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ae216-159">Type</span><span class="sxs-lookup"><span data-stu-id="ae216-159">Type</span></span>

*   <span data-ttu-id="ae216-160">String</span><span class="sxs-lookup"><span data-stu-id="ae216-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae216-161">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ae216-161">Requirements</span></span>

|<span data-ttu-id="ae216-162">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-162">Requirement</span></span>| <span data-ttu-id="ae216-163">Valeur</span><span class="sxs-lookup"><span data-stu-id="ae216-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae216-164">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae216-165">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-165">1.1</span></span>|
|[<span data-ttu-id="ae216-166">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ae216-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae216-167">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae216-168">Exemple</span><span class="sxs-lookup"><span data-stu-id="ae216-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ae216-169">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ae216-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ae216-170">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ae216-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ae216-171">Type</span><span class="sxs-lookup"><span data-stu-id="ae216-171">Type</span></span>

*   [<span data-ttu-id="ae216-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ae216-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ae216-173">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ae216-173">Requirements</span></span>

|<span data-ttu-id="ae216-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-174">Requirement</span></span>| <span data-ttu-id="ae216-175">Valeur</span><span class="sxs-lookup"><span data-stu-id="ae216-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae216-176">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae216-177">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-177">1.1</span></span>|
|[<span data-ttu-id="ae216-178">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ae216-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae216-179">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae216-180">Exemple</span><span class="sxs-lookup"><span data-stu-id="ae216-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ae216-181">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="ae216-181">displayLanguage: String</span></span>

<span data-ttu-id="ae216-182">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="ae216-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="ae216-183">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="ae216-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ae216-184">Type</span><span class="sxs-lookup"><span data-stu-id="ae216-184">Type</span></span>

*   <span data-ttu-id="ae216-185">String</span><span class="sxs-lookup"><span data-stu-id="ae216-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ae216-186">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ae216-186">Requirements</span></span>

|<span data-ttu-id="ae216-187">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-187">Requirement</span></span>| <span data-ttu-id="ae216-188">Valeur</span><span class="sxs-lookup"><span data-stu-id="ae216-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae216-189">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae216-190">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-190">1.1</span></span>|
|[<span data-ttu-id="ae216-191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ae216-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae216-192">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae216-193">Exemple</span><span class="sxs-lookup"><span data-stu-id="ae216-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ae216-194">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ae216-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ae216-195">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="ae216-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ae216-196">Type</span><span class="sxs-lookup"><span data-stu-id="ae216-196">Type</span></span>

*   [<span data-ttu-id="ae216-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ae216-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ae216-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ae216-198">Requirements</span></span>

|<span data-ttu-id="ae216-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-199">Requirement</span></span>| <span data-ttu-id="ae216-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="ae216-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae216-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae216-202">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-202">1.1</span></span>|
|[<span data-ttu-id="ae216-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ae216-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae216-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ae216-205">Exemple</span><span class="sxs-lookup"><span data-stu-id="ae216-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ae216-206">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ae216-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ae216-207">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ae216-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ae216-208">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="ae216-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ae216-209">Type</span><span class="sxs-lookup"><span data-stu-id="ae216-209">Type</span></span>

*   [<span data-ttu-id="ae216-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ae216-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ae216-211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ae216-211">Requirements</span></span>

|<span data-ttu-id="ae216-212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-212">Requirement</span></span>| <span data-ttu-id="ae216-213">Valeur</span><span class="sxs-lookup"><span data-stu-id="ae216-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae216-214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae216-215">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-215">1.1</span></span>|
|[<span data-ttu-id="ae216-216">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ae216-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ae216-217">Restreinte</span><span class="sxs-lookup"><span data-stu-id="ae216-217">Restricted</span></span>|
|[<span data-ttu-id="ae216-218">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ae216-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae216-219">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ae216-220">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ae216-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ae216-221">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="ae216-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ae216-222">Type</span><span class="sxs-lookup"><span data-stu-id="ae216-222">Type</span></span>

*   [<span data-ttu-id="ae216-223">UI</span><span class="sxs-lookup"><span data-stu-id="ae216-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ae216-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ae216-224">Requirements</span></span>

|<span data-ttu-id="ae216-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ae216-225">Requirement</span></span>| <span data-ttu-id="ae216-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="ae216-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="ae216-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ae216-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ae216-228">1.1</span><span class="sxs-lookup"><span data-stu-id="ae216-228">1.1</span></span>|
|[<span data-ttu-id="ae216-229">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ae216-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ae216-230">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ae216-230">Compose or Read</span></span>|
