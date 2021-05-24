---
title: Office.context - ensemble de conditions requises 1.10
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.10.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: cb189dc3b7b51357dee8ac83bc61795b3ec47ae5
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592038"
---
# <a name="context-mailbox-requirement-set-110"></a><span data-ttu-id="9b613-103">context (Ensemble de conditions requises de boîte aux lettres 1.10)</span><span class="sxs-lookup"><span data-stu-id="9b613-103">context (Mailbox requirement set 1.10)</span></span>

### <a name="officecontext"></a><span data-ttu-id="9b613-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="9b613-104">[Office](office.md).context</span></span>

<span data-ttu-id="9b613-105">Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications.</span><span class="sxs-lookup"><span data-stu-id="9b613-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9b613-106">Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir la liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="9b613-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b613-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-107">Requirements</span></span>

|<span data-ttu-id="9b613-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-108">Requirement</span></span>| <span data-ttu-id="9b613-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-111">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-111">1.1</span></span>|
|[<span data-ttu-id="9b613-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="9b613-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="9b613-114">Properties</span></span>

| <span data-ttu-id="9b613-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="9b613-115">Property</span></span> | <span data-ttu-id="9b613-116">Modes</span><span class="sxs-lookup"><span data-stu-id="9b613-116">Modes</span></span> | <span data-ttu-id="9b613-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="9b613-117">Return type</span></span> | <span data-ttu-id="9b613-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="9b613-118">Minimum</span></span><br><span data-ttu-id="9b613-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9b613-120">auth</span><span class="sxs-lookup"><span data-stu-id="9b613-120">auth</span></span>](#auth-auth) | <span data-ttu-id="9b613-121">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-121">Compose</span></span><br><span data-ttu-id="9b613-122">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-122">Read</span></span> | [<span data-ttu-id="9b613-123">Auth</span><span class="sxs-lookup"><span data-stu-id="9b613-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="9b613-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="9b613-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="9b613-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="9b613-126">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-126">Compose</span></span><br><span data-ttu-id="9b613-127">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-127">Read</span></span> | <span data-ttu-id="9b613-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9b613-128">String</span></span> | [<span data-ttu-id="9b613-129">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b613-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="9b613-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="9b613-131">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-131">Compose</span></span><br><span data-ttu-id="9b613-132">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-132">Read</span></span> | [<span data-ttu-id="9b613-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9b613-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b613-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="9b613-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="9b613-136">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-136">Compose</span></span><br><span data-ttu-id="9b613-137">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-137">Read</span></span> | <span data-ttu-id="9b613-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="9b613-138">String</span></span> | [<span data-ttu-id="9b613-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b613-140">host</span><span class="sxs-lookup"><span data-stu-id="9b613-140">host</span></span>](#host-hosttype) | <span data-ttu-id="9b613-141">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-141">Compose</span></span><br><span data-ttu-id="9b613-142">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-142">Read</span></span> | [<span data-ttu-id="9b613-143">HostType</span><span class="sxs-lookup"><span data-stu-id="9b613-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-144">1.5</span><span class="sxs-lookup"><span data-stu-id="9b613-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9b613-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="9b613-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="9b613-146">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-146">Compose</span></span><br><span data-ttu-id="9b613-147">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-147">Read</span></span> | [<span data-ttu-id="9b613-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b613-150">platform</span><span class="sxs-lookup"><span data-stu-id="9b613-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="9b613-151">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-151">Compose</span></span><br><span data-ttu-id="9b613-152">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-152">Read</span></span> | [<span data-ttu-id="9b613-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9b613-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-154">1.5</span><span class="sxs-lookup"><span data-stu-id="9b613-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9b613-155">requirements</span><span class="sxs-lookup"><span data-stu-id="9b613-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="9b613-156">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-156">Compose</span></span><br><span data-ttu-id="9b613-157">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-157">Read</span></span> | [<span data-ttu-id="9b613-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9b613-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b613-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="9b613-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="9b613-161">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-161">Compose</span></span><br><span data-ttu-id="9b613-162">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-162">Read</span></span> | [<span data-ttu-id="9b613-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9b613-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9b613-165">ui</span><span class="sxs-lookup"><span data-stu-id="9b613-165">ui</span></span>](#ui-ui) | <span data-ttu-id="9b613-166">Composition</span><span class="sxs-lookup"><span data-stu-id="9b613-166">Compose</span></span><br><span data-ttu-id="9b613-167">Lire</span><span class="sxs-lookup"><span data-stu-id="9b613-167">Read</span></span> | [<span data-ttu-id="9b613-168">UI</span><span class="sxs-lookup"><span data-stu-id="9b613-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9b613-169">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="9b613-170">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="9b613-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="9b613-171">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="9b613-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="9b613-172">Prend en charge l' [sign-on unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application web du module.</span><span class="sxs-lookup"><span data-stu-id="9b613-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="9b613-173">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="9b613-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-174">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-174">Type</span></span>

*   [<span data-ttu-id="9b613-175">Auth</span><span class="sxs-lookup"><span data-stu-id="9b613-175">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="9b613-176">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-176">Requirements</span></span>

|<span data-ttu-id="9b613-177">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-177">Requirement</span></span>| <span data-ttu-id="9b613-178">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-178">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-179">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-179">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-180">1.10</span><span class="sxs-lookup"><span data-stu-id="9b613-180">1.10</span></span>|
|[<span data-ttu-id="9b613-181">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-181">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-182">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-182">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b613-183">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b613-183">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="9b613-184">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="9b613-184">contentLanguage: String</span></span>

<span data-ttu-id="9b613-185">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="9b613-185">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="9b613-186">La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options d'> langue dans `contentLanguage` l’application cliente Office’édition.  </span><span class="sxs-lookup"><span data-stu-id="9b613-186">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-187">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-187">Type</span></span>

*   <span data-ttu-id="9b613-188">String</span><span class="sxs-lookup"><span data-stu-id="9b613-188">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b613-189">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-189">Requirements</span></span>

|<span data-ttu-id="9b613-190">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-190">Requirement</span></span>| <span data-ttu-id="9b613-191">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-191">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-192">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-192">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-193">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-193">1.1</span></span>|
|[<span data-ttu-id="9b613-194">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-194">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-195">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-195">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b613-196">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b613-196">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="9b613-197">diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9b613-197">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="9b613-198">Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="9b613-198">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-199">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-199">Type</span></span>

*   [<span data-ttu-id="9b613-200">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9b613-200">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="9b613-201">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-201">Requirements</span></span>

|<span data-ttu-id="9b613-202">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-202">Requirement</span></span>| <span data-ttu-id="9b613-203">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-203">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-204">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-204">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-205">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-205">1.1</span></span>|
|[<span data-ttu-id="9b613-206">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-206">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-207">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-207">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b613-208">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b613-208">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="9b613-209">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="9b613-209">displayLanguage: String</span></span>

<span data-ttu-id="9b613-210">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.</span><span class="sxs-lookup"><span data-stu-id="9b613-210">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="9b613-211">La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office..  </span><span class="sxs-lookup"><span data-stu-id="9b613-211">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-212">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-212">Type</span></span>

*   <span data-ttu-id="9b613-213">String</span><span class="sxs-lookup"><span data-stu-id="9b613-213">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9b613-214">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-214">Requirements</span></span>

|<span data-ttu-id="9b613-215">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-215">Requirement</span></span>| <span data-ttu-id="9b613-216">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-217">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-217">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-218">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-218">1.1</span></span>|
|[<span data-ttu-id="9b613-219">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-219">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-220">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b613-221">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b613-221">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="9b613-222">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="9b613-222">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="9b613-223">Obtient Office application qui héberge le module.</span><span class="sxs-lookup"><span data-stu-id="9b613-223">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9b613-224">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="9b613-224">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-225">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-225">Type</span></span>

*   [<span data-ttu-id="9b613-226">HostType</span><span class="sxs-lookup"><span data-stu-id="9b613-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="9b613-227">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-227">Requirements</span></span>

|<span data-ttu-id="9b613-228">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-228">Requirement</span></span>| <span data-ttu-id="9b613-229">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-230">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-231">1,5</span><span class="sxs-lookup"><span data-stu-id="9b613-231">1.5</span></span>|
|[<span data-ttu-id="9b613-232">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-233">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b613-234">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b613-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="9b613-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="9b613-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="9b613-236">Fournit la plateforme sur laquelle le module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="9b613-236">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="9b613-237">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="9b613-237">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-238">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-238">Type</span></span>

*   [<span data-ttu-id="9b613-239">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9b613-239">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="9b613-240">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-240">Requirements</span></span>

|<span data-ttu-id="9b613-241">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-241">Requirement</span></span>| <span data-ttu-id="9b613-242">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-243">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-244">1,5</span><span class="sxs-lookup"><span data-stu-id="9b613-244">1.5</span></span>|
|[<span data-ttu-id="9b613-245">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-246">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b613-247">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b613-247">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="9b613-248">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="9b613-248">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="9b613-249">Fournit une méthode pour déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="9b613-249">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-250">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-250">Type</span></span>

*   [<span data-ttu-id="9b613-251">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9b613-251">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="9b613-252">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-252">Requirements</span></span>

|<span data-ttu-id="9b613-253">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-253">Requirement</span></span>| <span data-ttu-id="9b613-254">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-255">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-255">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-256">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-256">1.1</span></span>|
|[<span data-ttu-id="9b613-257">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-257">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-258">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9b613-259">Exemple</span><span class="sxs-lookup"><span data-stu-id="9b613-259">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="9b613-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="9b613-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="9b613-261">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9b613-261">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9b613-262">L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.</span><span class="sxs-lookup"><span data-stu-id="9b613-262">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-263">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-263">Type</span></span>

*   [<span data-ttu-id="9b613-264">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9b613-264">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9b613-265">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-265">Requirements</span></span>

|<span data-ttu-id="9b613-266">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-266">Requirement</span></span>| <span data-ttu-id="9b613-267">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-268">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-268">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-269">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-269">1.1</span></span>|
|[<span data-ttu-id="9b613-270">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="9b613-270">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="9b613-271">Restreinte</span><span class="sxs-lookup"><span data-stu-id="9b613-271">Restricted</span></span>|
|[<span data-ttu-id="9b613-272">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-272">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-273">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="9b613-274">Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="9b613-274">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="9b613-275">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.</span><span class="sxs-lookup"><span data-stu-id="9b613-275">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9b613-276">Type</span><span class="sxs-lookup"><span data-stu-id="9b613-276">Type</span></span>

*   [<span data-ttu-id="9b613-277">UI</span><span class="sxs-lookup"><span data-stu-id="9b613-277">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="9b613-278">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="9b613-278">Requirements</span></span>

|<span data-ttu-id="9b613-279">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="9b613-279">Requirement</span></span>| <span data-ttu-id="9b613-280">Valeur</span><span class="sxs-lookup"><span data-stu-id="9b613-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="9b613-281">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="9b613-281">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9b613-282">1.1</span><span class="sxs-lookup"><span data-stu-id="9b613-282">1.1</span></span>|
|[<span data-ttu-id="9b613-283">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="9b613-283">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9b613-284">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="9b613-284">Compose or Read</span></span>|
