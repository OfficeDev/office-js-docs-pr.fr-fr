---
title: 'Office.context : ensemble de conditions requises de prévisualisation'
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 59b1cce579afe69384e41a6f31cc70c8cec25bea
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591071"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="d7048-103">context (ensemble de conditions requises d’aperçu de boîte aux lettres)</span><span class="sxs-lookup"><span data-stu-id="d7048-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="d7048-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="d7048-104">[Office](office.md).context</span></span>

<span data-ttu-id="d7048-105">Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications.</span><span class="sxs-lookup"><span data-stu-id="d7048-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="d7048-106">Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir la liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="d7048-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7048-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-107">Requirements</span></span>

|<span data-ttu-id="d7048-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-108">Requirement</span></span>| <span data-ttu-id="d7048-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-111">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-111">1.1</span></span>|
|[<span data-ttu-id="d7048-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="d7048-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="d7048-114">Properties</span></span>

| <span data-ttu-id="d7048-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="d7048-115">Property</span></span> | <span data-ttu-id="d7048-116">Modes</span><span class="sxs-lookup"><span data-stu-id="d7048-116">Modes</span></span> | <span data-ttu-id="d7048-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="d7048-117">Return type</span></span> | <span data-ttu-id="d7048-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="d7048-118">Minimum</span></span><br><span data-ttu-id="d7048-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d7048-120">auth</span><span class="sxs-lookup"><span data-stu-id="d7048-120">auth</span></span>](#auth-auth) | <span data-ttu-id="d7048-121">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-121">Compose</span></span><br><span data-ttu-id="d7048-122">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-122">Read</span></span> | [<span data-ttu-id="d7048-123">Auth</span><span class="sxs-lookup"><span data-stu-id="d7048-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="d7048-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="d7048-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="d7048-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="d7048-126">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-126">Compose</span></span><br><span data-ttu-id="d7048-127">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-127">Read</span></span> | <span data-ttu-id="d7048-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7048-128">String</span></span> | [<span data-ttu-id="d7048-129">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7048-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="d7048-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="d7048-131">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-131">Compose</span></span><br><span data-ttu-id="d7048-132">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-132">Read</span></span> | [<span data-ttu-id="d7048-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="d7048-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7048-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="d7048-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="d7048-136">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-136">Compose</span></span><br><span data-ttu-id="d7048-137">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-137">Read</span></span> | <span data-ttu-id="d7048-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7048-138">String</span></span> | [<span data-ttu-id="d7048-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7048-140">host</span><span class="sxs-lookup"><span data-stu-id="d7048-140">host</span></span>](#host-hosttype) | <span data-ttu-id="d7048-141">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-141">Compose</span></span><br><span data-ttu-id="d7048-142">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-142">Read</span></span> | [<span data-ttu-id="d7048-143">HostType</span><span class="sxs-lookup"><span data-stu-id="d7048-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-144">1.5</span><span class="sxs-lookup"><span data-stu-id="d7048-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="d7048-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="d7048-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="d7048-146">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-146">Compose</span></span><br><span data-ttu-id="d7048-147">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-147">Read</span></span> | [<span data-ttu-id="d7048-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7048-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="d7048-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="d7048-151">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-151">Compose</span></span><br><span data-ttu-id="d7048-152">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-152">Read</span></span> | [<span data-ttu-id="d7048-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d7048-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-154">Aperçu</span><span class="sxs-lookup"><span data-stu-id="d7048-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="d7048-155">platform</span><span class="sxs-lookup"><span data-stu-id="d7048-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="d7048-156">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-156">Compose</span></span><br><span data-ttu-id="d7048-157">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-157">Read</span></span> | [<span data-ttu-id="d7048-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d7048-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-159">1.5</span><span class="sxs-lookup"><span data-stu-id="d7048-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="d7048-160">requirements</span><span class="sxs-lookup"><span data-stu-id="d7048-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="d7048-161">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-161">Compose</span></span><br><span data-ttu-id="d7048-162">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-162">Read</span></span> | [<span data-ttu-id="d7048-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="d7048-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-164">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7048-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="d7048-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="d7048-166">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-166">Compose</span></span><br><span data-ttu-id="d7048-167">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-167">Read</span></span> | [<span data-ttu-id="d7048-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d7048-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d7048-170">ui</span><span class="sxs-lookup"><span data-stu-id="d7048-170">ui</span></span>](#ui-ui) | <span data-ttu-id="d7048-171">Composition</span><span class="sxs-lookup"><span data-stu-id="d7048-171">Compose</span></span><br><span data-ttu-id="d7048-172">Lire</span><span class="sxs-lookup"><span data-stu-id="d7048-172">Read</span></span> | [<span data-ttu-id="d7048-173">UI</span><span class="sxs-lookup"><span data-stu-id="d7048-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="d7048-174">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="d7048-175">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="d7048-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="d7048-176">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="d7048-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="d7048-177">Prend en charge l' [sign-on unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application web du module.</span><span class="sxs-lookup"><span data-stu-id="d7048-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="d7048-178">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="d7048-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-179">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-179">Type</span></span>

*   [<span data-ttu-id="d7048-180">Auth</span><span class="sxs-lookup"><span data-stu-id="d7048-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="d7048-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-181">Requirements</span></span>

|<span data-ttu-id="d7048-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-182">Requirement</span></span>| <span data-ttu-id="d7048-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-185">Aperçu</span><span class="sxs-lookup"><span data-stu-id="d7048-185">Preview</span></span>|
|[<span data-ttu-id="d7048-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="d7048-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="d7048-189">contentLanguage: String</span></span>

<span data-ttu-id="d7048-190">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="d7048-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="d7048-191">La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options de > langue dans l Office `contentLanguage` application cliente.  </span><span class="sxs-lookup"><span data-stu-id="d7048-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-192">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-192">Type</span></span>

*   <span data-ttu-id="d7048-193">String</span><span class="sxs-lookup"><span data-stu-id="d7048-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7048-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-194">Requirements</span></span>

|<span data-ttu-id="d7048-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-195">Requirement</span></span>| <span data-ttu-id="d7048-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-198">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-198">1.1</span></span>|
|[<span data-ttu-id="d7048-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-201">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="d7048-202">diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="d7048-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="d7048-203">Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="d7048-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-204">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-204">Type</span></span>

*   [<span data-ttu-id="d7048-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="d7048-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="d7048-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-206">Requirements</span></span>

|<span data-ttu-id="d7048-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-207">Requirement</span></span>| <span data-ttu-id="d7048-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-210">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-210">1.1</span></span>|
|[<span data-ttu-id="d7048-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-212">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="d7048-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="d7048-214">displayLanguage: String</span></span>

<span data-ttu-id="d7048-215">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.</span><span class="sxs-lookup"><span data-stu-id="d7048-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="d7048-216">La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office..  </span><span class="sxs-lookup"><span data-stu-id="d7048-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-217">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-217">Type</span></span>

*   <span data-ttu-id="d7048-218">String</span><span class="sxs-lookup"><span data-stu-id="d7048-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7048-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-219">Requirements</span></span>

|<span data-ttu-id="d7048-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-220">Requirement</span></span>| <span data-ttu-id="d7048-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-223">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-223">1.1</span></span>|
|[<span data-ttu-id="d7048-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="d7048-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="d7048-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="d7048-228">Obtient Office application qui héberge le module.</span><span class="sxs-lookup"><span data-stu-id="d7048-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="d7048-229">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="d7048-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-230">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-230">Type</span></span>

*   [<span data-ttu-id="d7048-231">HostType</span><span class="sxs-lookup"><span data-stu-id="d7048-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="d7048-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-232">Requirements</span></span>

|<span data-ttu-id="d7048-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-233">Requirement</span></span>| <span data-ttu-id="d7048-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-236">1,5</span><span class="sxs-lookup"><span data-stu-id="d7048-236">1.5</span></span>|
|[<span data-ttu-id="d7048-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="d7048-240">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="d7048-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="d7048-241">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="d7048-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="d7048-242">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="d7048-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="d7048-243">L’utilisation de couleurs de thème Office vous permet de coordonner le modèle de couleurs de votre application avec le thème Office actuel sélectionné par l’utilisateur avec l’interface utilisateur de thème du compte **> Office** de > Office, qui est appliquée à toutes les applications clientes Office.</span><span class="sxs-lookup"><span data-stu-id="d7048-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="d7048-244">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="d7048-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-245">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-245">Type</span></span>

*   [<span data-ttu-id="d7048-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d7048-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="d7048-247">Propriétés</span><span class="sxs-lookup"><span data-stu-id="d7048-247">Properties</span></span>

|<span data-ttu-id="d7048-248">Nom</span><span class="sxs-lookup"><span data-stu-id="d7048-248">Name</span></span>| <span data-ttu-id="d7048-249">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-249">Type</span></span>| <span data-ttu-id="d7048-250">Description</span><span class="sxs-lookup"><span data-stu-id="d7048-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="d7048-251">Chaîne</span><span class="sxs-lookup"><span data-stu-id="d7048-251">String</span></span>|<span data-ttu-id="d7048-252">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="d7048-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="d7048-253">String</span><span class="sxs-lookup"><span data-stu-id="d7048-253">String</span></span>|<span data-ttu-id="d7048-254">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="d7048-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="d7048-255">String</span><span class="sxs-lookup"><span data-stu-id="d7048-255">String</span></span>|<span data-ttu-id="d7048-256">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="d7048-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="d7048-257">String</span><span class="sxs-lookup"><span data-stu-id="d7048-257">String</span></span>|<span data-ttu-id="d7048-258">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="d7048-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7048-259">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-259">Requirements</span></span>

|<span data-ttu-id="d7048-260">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-260">Requirement</span></span>| <span data-ttu-id="d7048-261">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-262">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-263">Aperçu</span><span class="sxs-lookup"><span data-stu-id="d7048-263">Preview</span></span>|
|[<span data-ttu-id="d7048-264">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-265">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-266">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-266">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="d7048-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="d7048-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="d7048-268">Fournit la plateforme sur laquelle le module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="d7048-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="d7048-269">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="d7048-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-270">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-270">Type</span></span>

*   [<span data-ttu-id="d7048-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d7048-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="d7048-272">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-272">Requirements</span></span>

|<span data-ttu-id="d7048-273">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-273">Requirement</span></span>| <span data-ttu-id="d7048-274">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-275">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-276">1,5</span><span class="sxs-lookup"><span data-stu-id="d7048-276">1.5</span></span>|
|[<span data-ttu-id="d7048-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-278">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-279">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="d7048-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="d7048-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="d7048-281">Fournit une méthode pour déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="d7048-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-282">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-282">Type</span></span>

*   [<span data-ttu-id="d7048-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="d7048-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="d7048-284">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-284">Requirements</span></span>

|<span data-ttu-id="d7048-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-285">Requirement</span></span>| <span data-ttu-id="d7048-286">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-287">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-288">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-288">1.1</span></span>|
|[<span data-ttu-id="d7048-289">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-290">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d7048-291">Exemple</span><span class="sxs-lookup"><span data-stu-id="d7048-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="d7048-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="d7048-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="d7048-293">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d7048-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="d7048-294">L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.</span><span class="sxs-lookup"><span data-stu-id="d7048-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-295">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-295">Type</span></span>

*   [<span data-ttu-id="d7048-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d7048-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="d7048-297">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-297">Requirements</span></span>

|<span data-ttu-id="d7048-298">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-298">Requirement</span></span>| <span data-ttu-id="d7048-299">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-300">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-301">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-301">1.1</span></span>|
|[<span data-ttu-id="d7048-302">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="d7048-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="d7048-303">Restreinte</span><span class="sxs-lookup"><span data-stu-id="d7048-303">Restricted</span></span>|
|[<span data-ttu-id="d7048-304">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-305">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="d7048-306">Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="d7048-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="d7048-307">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.</span><span class="sxs-lookup"><span data-stu-id="d7048-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d7048-308">Type</span><span class="sxs-lookup"><span data-stu-id="d7048-308">Type</span></span>

*   [<span data-ttu-id="d7048-309">UI</span><span class="sxs-lookup"><span data-stu-id="d7048-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="d7048-310">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="d7048-310">Requirements</span></span>

|<span data-ttu-id="d7048-311">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="d7048-311">Requirement</span></span>| <span data-ttu-id="d7048-312">Valeur</span><span class="sxs-lookup"><span data-stu-id="d7048-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7048-313">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="d7048-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d7048-314">1.1</span><span class="sxs-lookup"><span data-stu-id="d7048-314">1.1</span></span>|
|[<span data-ttu-id="d7048-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="d7048-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d7048-316">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="d7048-316">Compose or Read</span></span>|
