---
title: Office.context - ensemble de conditions requises 1.1
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.1.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 41273bfc5362a9d5572e38b8e80b81041f5aa312
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590875"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="cafef-103">context (Ensemble de conditions requises de boîte aux lettres 1.1)</span><span class="sxs-lookup"><span data-stu-id="cafef-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="cafef-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="cafef-104">[Office](office.md).context</span></span>

<span data-ttu-id="cafef-105">Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications.</span><span class="sxs-lookup"><span data-stu-id="cafef-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="cafef-106">Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir la liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="cafef-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="cafef-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cafef-107">Requirements</span></span>

|<span data-ttu-id="cafef-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-108">Requirement</span></span>| <span data-ttu-id="cafef-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="cafef-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="cafef-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cafef-111">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-111">1.1</span></span>|
|[<span data-ttu-id="cafef-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cafef-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cafef-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cafef-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="cafef-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="cafef-114">Properties</span></span>

| <span data-ttu-id="cafef-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="cafef-115">Property</span></span> | <span data-ttu-id="cafef-116">Modes</span><span class="sxs-lookup"><span data-stu-id="cafef-116">Modes</span></span> | <span data-ttu-id="cafef-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="cafef-117">Return type</span></span> | <span data-ttu-id="cafef-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="cafef-118">Minimum</span></span><br><span data-ttu-id="cafef-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="cafef-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="cafef-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="cafef-121">Composition</span><span class="sxs-lookup"><span data-stu-id="cafef-121">Compose</span></span><br><span data-ttu-id="cafef-122">Lire</span><span class="sxs-lookup"><span data-stu-id="cafef-122">Read</span></span> | <span data-ttu-id="cafef-123">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cafef-123">String</span></span> | [<span data-ttu-id="cafef-124">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cafef-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="cafef-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="cafef-126">Composition</span><span class="sxs-lookup"><span data-stu-id="cafef-126">Compose</span></span><br><span data-ttu-id="cafef-127">Lire</span><span class="sxs-lookup"><span data-stu-id="cafef-127">Read</span></span> | [<span data-ttu-id="cafef-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="cafef-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="cafef-129">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cafef-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="cafef-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="cafef-131">Composition</span><span class="sxs-lookup"><span data-stu-id="cafef-131">Compose</span></span><br><span data-ttu-id="cafef-132">Lire</span><span class="sxs-lookup"><span data-stu-id="cafef-132">Read</span></span> | <span data-ttu-id="cafef-133">Chaîne</span><span class="sxs-lookup"><span data-stu-id="cafef-133">String</span></span> | [<span data-ttu-id="cafef-134">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cafef-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="cafef-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="cafef-136">Composition</span><span class="sxs-lookup"><span data-stu-id="cafef-136">Compose</span></span><br><span data-ttu-id="cafef-137">Lire</span><span class="sxs-lookup"><span data-stu-id="cafef-137">Read</span></span> | [<span data-ttu-id="cafef-138">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="cafef-139">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cafef-140">requirements</span><span class="sxs-lookup"><span data-stu-id="cafef-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="cafef-141">Composition</span><span class="sxs-lookup"><span data-stu-id="cafef-141">Compose</span></span><br><span data-ttu-id="cafef-142">Lire</span><span class="sxs-lookup"><span data-stu-id="cafef-142">Read</span></span> | [<span data-ttu-id="cafef-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="cafef-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="cafef-144">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cafef-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="cafef-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="cafef-146">Composition</span><span class="sxs-lookup"><span data-stu-id="cafef-146">Compose</span></span><br><span data-ttu-id="cafef-147">Lire</span><span class="sxs-lookup"><span data-stu-id="cafef-147">Read</span></span> | [<span data-ttu-id="cafef-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cafef-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="cafef-149">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="cafef-150">ui</span><span class="sxs-lookup"><span data-stu-id="cafef-150">ui</span></span>](#ui-ui) | <span data-ttu-id="cafef-151">Composition</span><span class="sxs-lookup"><span data-stu-id="cafef-151">Compose</span></span><br><span data-ttu-id="cafef-152">Lire</span><span class="sxs-lookup"><span data-stu-id="cafef-152">Read</span></span> | [<span data-ttu-id="cafef-153">UI</span><span class="sxs-lookup"><span data-stu-id="cafef-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="cafef-154">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="cafef-155">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="cafef-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="cafef-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="cafef-156">contentLanguage: String</span></span>

<span data-ttu-id="cafef-157">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="cafef-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="cafef-158">La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options d'> langue dans `contentLanguage` l’application cliente Office’édition.  </span><span class="sxs-lookup"><span data-stu-id="cafef-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="cafef-159">Type</span><span class="sxs-lookup"><span data-stu-id="cafef-159">Type</span></span>

*   <span data-ttu-id="cafef-160">String</span><span class="sxs-lookup"><span data-stu-id="cafef-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cafef-161">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cafef-161">Requirements</span></span>

|<span data-ttu-id="cafef-162">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-162">Requirement</span></span>| <span data-ttu-id="cafef-163">Valeur</span><span class="sxs-lookup"><span data-stu-id="cafef-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="cafef-164">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cafef-165">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-165">1.1</span></span>|
|[<span data-ttu-id="cafef-166">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cafef-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cafef-167">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cafef-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cafef-168">Exemple</span><span class="sxs-lookup"><span data-stu-id="cafef-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="cafef-169">diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="cafef-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="cafef-170">Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="cafef-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="cafef-171">Type</span><span class="sxs-lookup"><span data-stu-id="cafef-171">Type</span></span>

*   [<span data-ttu-id="cafef-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="cafef-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="cafef-173">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cafef-173">Requirements</span></span>

|<span data-ttu-id="cafef-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-174">Requirement</span></span>| <span data-ttu-id="cafef-175">Valeur</span><span class="sxs-lookup"><span data-stu-id="cafef-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="cafef-176">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cafef-177">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-177">1.1</span></span>|
|[<span data-ttu-id="cafef-178">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cafef-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cafef-179">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cafef-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cafef-180">Exemple</span><span class="sxs-lookup"><span data-stu-id="cafef-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="cafef-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="cafef-181">displayLanguage: String</span></span>

<span data-ttu-id="cafef-182">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.</span><span class="sxs-lookup"><span data-stu-id="cafef-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="cafef-183">La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office..  </span><span class="sxs-lookup"><span data-stu-id="cafef-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="cafef-184">Type</span><span class="sxs-lookup"><span data-stu-id="cafef-184">Type</span></span>

*   <span data-ttu-id="cafef-185">String</span><span class="sxs-lookup"><span data-stu-id="cafef-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cafef-186">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cafef-186">Requirements</span></span>

|<span data-ttu-id="cafef-187">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-187">Requirement</span></span>| <span data-ttu-id="cafef-188">Valeur</span><span class="sxs-lookup"><span data-stu-id="cafef-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="cafef-189">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cafef-190">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-190">1.1</span></span>|
|[<span data-ttu-id="cafef-191">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cafef-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cafef-192">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cafef-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cafef-193">Exemple</span><span class="sxs-lookup"><span data-stu-id="cafef-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="cafef-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="cafef-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="cafef-195">Fournit une méthode permettant de déterminer les ensembles de conditions requises pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="cafef-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="cafef-196">Type</span><span class="sxs-lookup"><span data-stu-id="cafef-196">Type</span></span>

*   [<span data-ttu-id="cafef-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="cafef-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="cafef-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cafef-198">Requirements</span></span>

|<span data-ttu-id="cafef-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-199">Requirement</span></span>| <span data-ttu-id="cafef-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="cafef-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="cafef-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cafef-202">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-202">1.1</span></span>|
|[<span data-ttu-id="cafef-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cafef-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cafef-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cafef-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cafef-205">Exemple</span><span class="sxs-lookup"><span data-stu-id="cafef-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="cafef-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="cafef-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="cafef-207">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="cafef-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="cafef-208">L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.</span><span class="sxs-lookup"><span data-stu-id="cafef-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="cafef-209">Type</span><span class="sxs-lookup"><span data-stu-id="cafef-209">Type</span></span>

*   [<span data-ttu-id="cafef-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="cafef-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="cafef-211">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cafef-211">Requirements</span></span>

|<span data-ttu-id="cafef-212">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-212">Requirement</span></span>| <span data-ttu-id="cafef-213">Valeur</span><span class="sxs-lookup"><span data-stu-id="cafef-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="cafef-214">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cafef-215">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-215">1.1</span></span>|
|[<span data-ttu-id="cafef-216">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="cafef-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="cafef-217">Restreinte</span><span class="sxs-lookup"><span data-stu-id="cafef-217">Restricted</span></span>|
|[<span data-ttu-id="cafef-218">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cafef-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cafef-219">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cafef-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="cafef-220">Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="cafef-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="cafef-221">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.</span><span class="sxs-lookup"><span data-stu-id="cafef-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="cafef-222">Type</span><span class="sxs-lookup"><span data-stu-id="cafef-222">Type</span></span>

*   [<span data-ttu-id="cafef-223">UI</span><span class="sxs-lookup"><span data-stu-id="cafef-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="cafef-224">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="cafef-224">Requirements</span></span>

|<span data-ttu-id="cafef-225">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="cafef-225">Requirement</span></span>| <span data-ttu-id="cafef-226">Valeur</span><span class="sxs-lookup"><span data-stu-id="cafef-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="cafef-227">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="cafef-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="cafef-228">1.1</span><span class="sxs-lookup"><span data-stu-id="cafef-228">1.1</span></span>|
|[<span data-ttu-id="cafef-229">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="cafef-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="cafef-230">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="cafef-230">Compose or Read</span></span>|
