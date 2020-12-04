---
title: Office. Context-ensemble de conditions requises 1,1
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,1.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 4f0fa4094477125f4d07fd6ddb4ac2c3c08a5d70
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570750"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="dbf32-103">contexte (boîte aux lettres requise définie sur 1,1)</span><span class="sxs-lookup"><span data-stu-id="dbf32-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="dbf32-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="dbf32-104">[Office](office.md).context</span></span>

<span data-ttu-id="dbf32-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="dbf32-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="dbf32-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="dbf32-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="dbf32-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-107">Requirements</span></span>

|<span data-ttu-id="dbf32-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-108">Requirement</span></span>| <span data-ttu-id="dbf32-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="dbf32-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="dbf32-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dbf32-111">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-111">1.1</span></span>|
|[<span data-ttu-id="dbf32-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dbf32-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dbf32-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="dbf32-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="dbf32-114">Properties</span></span>

| <span data-ttu-id="dbf32-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="dbf32-115">Property</span></span> | <span data-ttu-id="dbf32-116">Modes</span><span class="sxs-lookup"><span data-stu-id="dbf32-116">Modes</span></span> | <span data-ttu-id="dbf32-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="dbf32-117">Return type</span></span> | <span data-ttu-id="dbf32-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="dbf32-118">Minimum</span></span><br><span data-ttu-id="dbf32-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="dbf32-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="dbf32-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="dbf32-121">Composition</span><span class="sxs-lookup"><span data-stu-id="dbf32-121">Compose</span></span><br><span data-ttu-id="dbf32-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-122">Read</span></span> | <span data-ttu-id="dbf32-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dbf32-123">String</span></span> | [<span data-ttu-id="dbf32-124">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dbf32-125">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="dbf32-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="dbf32-126">Composition</span><span class="sxs-lookup"><span data-stu-id="dbf32-126">Compose</span></span><br><span data-ttu-id="dbf32-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-127">Read</span></span> | [<span data-ttu-id="dbf32-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dbf32-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="dbf32-129">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dbf32-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="dbf32-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="dbf32-131">Composition</span><span class="sxs-lookup"><span data-stu-id="dbf32-131">Compose</span></span><br><span data-ttu-id="dbf32-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-132">Read</span></span> | <span data-ttu-id="dbf32-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="dbf32-133">String</span></span> | [<span data-ttu-id="dbf32-134">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dbf32-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="dbf32-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="dbf32-136">Composition</span><span class="sxs-lookup"><span data-stu-id="dbf32-136">Compose</span></span><br><span data-ttu-id="dbf32-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-137">Read</span></span> | [<span data-ttu-id="dbf32-138">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="dbf32-139">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dbf32-140">requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="dbf32-141">Composition</span><span class="sxs-lookup"><span data-stu-id="dbf32-141">Compose</span></span><br><span data-ttu-id="dbf32-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-142">Read</span></span> | [<span data-ttu-id="dbf32-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dbf32-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="dbf32-144">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dbf32-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="dbf32-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="dbf32-146">Composition</span><span class="sxs-lookup"><span data-stu-id="dbf32-146">Compose</span></span><br><span data-ttu-id="dbf32-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-147">Read</span></span> | [<span data-ttu-id="dbf32-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dbf32-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="dbf32-149">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="dbf32-150">ui</span><span class="sxs-lookup"><span data-stu-id="dbf32-150">ui</span></span>](#ui-ui) | <span data-ttu-id="dbf32-151">Composition</span><span class="sxs-lookup"><span data-stu-id="dbf32-151">Compose</span></span><br><span data-ttu-id="dbf32-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-152">Read</span></span> | [<span data-ttu-id="dbf32-153">UI</span><span class="sxs-lookup"><span data-stu-id="dbf32-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="dbf32-154">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="dbf32-155">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="dbf32-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="dbf32-156">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="dbf32-156">contentLanguage: String</span></span>

<span data-ttu-id="dbf32-157">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="dbf32-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="dbf32-158">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="dbf32-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dbf32-159">Type</span><span class="sxs-lookup"><span data-stu-id="dbf32-159">Type</span></span>

*   <span data-ttu-id="dbf32-160">String</span><span class="sxs-lookup"><span data-stu-id="dbf32-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dbf32-161">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-161">Requirements</span></span>

|<span data-ttu-id="dbf32-162">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-162">Requirement</span></span>| <span data-ttu-id="dbf32-163">Valeur</span><span class="sxs-lookup"><span data-stu-id="dbf32-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="dbf32-164">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dbf32-165">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-165">1.1</span></span>|
|[<span data-ttu-id="dbf32-166">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dbf32-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dbf32-167">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dbf32-168">Exemple</span><span class="sxs-lookup"><span data-stu-id="dbf32-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="dbf32-169">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="dbf32-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="dbf32-170">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="dbf32-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="dbf32-171">Type</span><span class="sxs-lookup"><span data-stu-id="dbf32-171">Type</span></span>

*   [<span data-ttu-id="dbf32-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="dbf32-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="dbf32-173">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-173">Requirements</span></span>

|<span data-ttu-id="dbf32-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-174">Requirement</span></span>| <span data-ttu-id="dbf32-175">Valeur</span><span class="sxs-lookup"><span data-stu-id="dbf32-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="dbf32-176">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dbf32-177">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-177">1.1</span></span>|
|[<span data-ttu-id="dbf32-178">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dbf32-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dbf32-179">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dbf32-180">Exemple</span><span class="sxs-lookup"><span data-stu-id="dbf32-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="dbf32-181">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="dbf32-181">displayLanguage: String</span></span>

<span data-ttu-id="dbf32-182">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="dbf32-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="dbf32-183">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="dbf32-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="dbf32-184">Type</span><span class="sxs-lookup"><span data-stu-id="dbf32-184">Type</span></span>

*   <span data-ttu-id="dbf32-185">String</span><span class="sxs-lookup"><span data-stu-id="dbf32-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="dbf32-186">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-186">Requirements</span></span>

|<span data-ttu-id="dbf32-187">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-187">Requirement</span></span>| <span data-ttu-id="dbf32-188">Valeur</span><span class="sxs-lookup"><span data-stu-id="dbf32-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="dbf32-189">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dbf32-190">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-190">1.1</span></span>|
|[<span data-ttu-id="dbf32-191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dbf32-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dbf32-192">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dbf32-193">Exemple</span><span class="sxs-lookup"><span data-stu-id="dbf32-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="dbf32-194">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="dbf32-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="dbf32-195">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="dbf32-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="dbf32-196">Type</span><span class="sxs-lookup"><span data-stu-id="dbf32-196">Type</span></span>

*   [<span data-ttu-id="dbf32-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="dbf32-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="dbf32-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-198">Requirements</span></span>

|<span data-ttu-id="dbf32-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-199">Requirement</span></span>| <span data-ttu-id="dbf32-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="dbf32-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="dbf32-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dbf32-202">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-202">1.1</span></span>|
|[<span data-ttu-id="dbf32-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dbf32-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dbf32-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="dbf32-205">Exemple</span><span class="sxs-lookup"><span data-stu-id="dbf32-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="dbf32-206">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="dbf32-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="dbf32-207">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="dbf32-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="dbf32-208">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="dbf32-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="dbf32-209">Type</span><span class="sxs-lookup"><span data-stu-id="dbf32-209">Type</span></span>

*   [<span data-ttu-id="dbf32-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="dbf32-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="dbf32-211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-211">Requirements</span></span>

|<span data-ttu-id="dbf32-212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-212">Requirement</span></span>| <span data-ttu-id="dbf32-213">Valeur</span><span class="sxs-lookup"><span data-stu-id="dbf32-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="dbf32-214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dbf32-215">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-215">1.1</span></span>|
|[<span data-ttu-id="dbf32-216">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="dbf32-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="dbf32-217">Restreinte</span><span class="sxs-lookup"><span data-stu-id="dbf32-217">Restricted</span></span>|
|[<span data-ttu-id="dbf32-218">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dbf32-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dbf32-219">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="dbf32-220">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="dbf32-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="dbf32-221">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="dbf32-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="dbf32-222">Type</span><span class="sxs-lookup"><span data-stu-id="dbf32-222">Type</span></span>

*   [<span data-ttu-id="dbf32-223">UI</span><span class="sxs-lookup"><span data-stu-id="dbf32-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="dbf32-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="dbf32-224">Requirements</span></span>

|<span data-ttu-id="dbf32-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="dbf32-225">Requirement</span></span>| <span data-ttu-id="dbf32-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="dbf32-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="dbf32-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="dbf32-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="dbf32-228">1.1</span><span class="sxs-lookup"><span data-stu-id="dbf32-228">1.1</span></span>|
|[<span data-ttu-id="dbf32-229">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="dbf32-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="dbf32-230">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="dbf32-230">Compose or Read</span></span>|
