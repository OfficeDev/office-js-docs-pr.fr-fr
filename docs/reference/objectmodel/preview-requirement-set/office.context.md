---
title: Ensemble de conditions requises pour Office. Context-preview
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook à l’aide de l’ensemble de conditions requises pour l’API de boîte aux lettres.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 8286434d2cbfc11cf0d16f8bd014b4760f0337ff
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626406"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="ed427-103">contexte (ensemble de conditions requises pour la boîte aux lettres)</span><span class="sxs-lookup"><span data-stu-id="ed427-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="ed427-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="ed427-104">[Office](office.md).context</span></span>

<span data-ttu-id="ed427-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="ed427-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="ed427-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="ed427-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-107">Requirements</span></span>

|<span data-ttu-id="ed427-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-108">Requirement</span></span>| <span data-ttu-id="ed427-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-111">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-111">1.1</span></span>|
|[<span data-ttu-id="ed427-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="ed427-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="ed427-114">Properties</span></span>

| <span data-ttu-id="ed427-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="ed427-115">Property</span></span> | <span data-ttu-id="ed427-116">Modes</span><span class="sxs-lookup"><span data-stu-id="ed427-116">Modes</span></span> | <span data-ttu-id="ed427-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="ed427-117">Return type</span></span> | <span data-ttu-id="ed427-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="ed427-118">Minimum</span></span><br><span data-ttu-id="ed427-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="ed427-120">auth</span><span class="sxs-lookup"><span data-stu-id="ed427-120">auth</span></span>](#auth-auth) | <span data-ttu-id="ed427-121">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-121">Compose</span></span><br><span data-ttu-id="ed427-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-122">Read</span></span> | [<span data-ttu-id="ed427-123">Auth</span><span class="sxs-lookup"><span data-stu-id="ed427-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-124">Ensembles 1,3</span><span class="sxs-lookup"><span data-stu-id="ed427-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="ed427-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="ed427-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="ed427-126">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-126">Compose</span></span><br><span data-ttu-id="ed427-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-127">Read</span></span> | <span data-ttu-id="ed427-128">String</span><span class="sxs-lookup"><span data-stu-id="ed427-128">String</span></span> | [<span data-ttu-id="ed427-129">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-130">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="ed427-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="ed427-131">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-131">Compose</span></span><br><span data-ttu-id="ed427-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-132">Read</span></span> | [<span data-ttu-id="ed427-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ed427-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-134">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="ed427-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="ed427-136">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-136">Compose</span></span><br><span data-ttu-id="ed427-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-137">Read</span></span> | <span data-ttu-id="ed427-138">String</span><span class="sxs-lookup"><span data-stu-id="ed427-138">String</span></span> | [<span data-ttu-id="ed427-139">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-140">hote</span><span class="sxs-lookup"><span data-stu-id="ed427-140">host</span></span>](#host-hosttype) | <span data-ttu-id="ed427-141">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-141">Compose</span></span><br><span data-ttu-id="ed427-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-142">Read</span></span> | [<span data-ttu-id="ed427-143">HostType</span><span class="sxs-lookup"><span data-stu-id="ed427-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-144">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="ed427-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="ed427-146">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-146">Compose</span></span><br><span data-ttu-id="ed427-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-147">Read</span></span> | [<span data-ttu-id="ed427-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-149">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="ed427-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="ed427-151">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-151">Compose</span></span><br><span data-ttu-id="ed427-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-152">Read</span></span> | [<span data-ttu-id="ed427-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="ed427-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-154">Aperçu</span><span class="sxs-lookup"><span data-stu-id="ed427-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="ed427-155">platform</span><span class="sxs-lookup"><span data-stu-id="ed427-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="ed427-156">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-156">Compose</span></span><br><span data-ttu-id="ed427-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-157">Read</span></span> | [<span data-ttu-id="ed427-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ed427-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-159">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-160">requise</span><span class="sxs-lookup"><span data-stu-id="ed427-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="ed427-161">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-161">Compose</span></span><br><span data-ttu-id="ed427-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-162">Read</span></span> | [<span data-ttu-id="ed427-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ed427-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-164">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="ed427-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="ed427-166">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-166">Compose</span></span><br><span data-ttu-id="ed427-167">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-167">Read</span></span> | [<span data-ttu-id="ed427-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ed427-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-169">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="ed427-170">ui</span><span class="sxs-lookup"><span data-stu-id="ed427-170">ui</span></span>](#ui-ui) | <span data-ttu-id="ed427-171">Composition</span><span class="sxs-lookup"><span data-stu-id="ed427-171">Compose</span></span><br><span data-ttu-id="ed427-172">Lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-172">Read</span></span> | [<span data-ttu-id="ed427-173">UI</span><span class="sxs-lookup"><span data-stu-id="ed427-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="ed427-174">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="ed427-175">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="ed427-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="ed427-176">AUTH : [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="ed427-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="ed427-177">Prend en charge l’authentification [unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application Web du complément.</span><span class="sxs-lookup"><span data-stu-id="ed427-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="ed427-178">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="ed427-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-179">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-179">Type</span></span>

*   [<span data-ttu-id="ed427-180">Auth</span><span class="sxs-lookup"><span data-stu-id="ed427-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="ed427-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-181">Requirements</span></span>

|<span data-ttu-id="ed427-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-182">Requirement</span></span>| <span data-ttu-id="ed427-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-185">Aperçu</span><span class="sxs-lookup"><span data-stu-id="ed427-185">Preview</span></span>|
|[<span data-ttu-id="ed427-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="ed427-189">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="ed427-189">contentLanguage: String</span></span>

<span data-ttu-id="ed427-190">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="ed427-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="ed427-191">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-192">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-192">Type</span></span>

*   <span data-ttu-id="ed427-193">String</span><span class="sxs-lookup"><span data-stu-id="ed427-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ed427-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-194">Requirements</span></span>

|<span data-ttu-id="ed427-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-195">Requirement</span></span>| <span data-ttu-id="ed427-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-198">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-198">1.1</span></span>|
|[<span data-ttu-id="ed427-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-201">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="ed427-202">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="ed427-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="ed427-203">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ed427-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-204">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-204">Type</span></span>

*   [<span data-ttu-id="ed427-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="ed427-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="ed427-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-206">Requirements</span></span>

|<span data-ttu-id="ed427-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-207">Requirement</span></span>| <span data-ttu-id="ed427-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-210">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-210">1.1</span></span>|
|[<span data-ttu-id="ed427-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-212">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="ed427-214">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="ed427-214">displayLanguage: String</span></span>

<span data-ttu-id="ed427-215">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="ed427-216">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-217">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-217">Type</span></span>

*   <span data-ttu-id="ed427-218">String</span><span class="sxs-lookup"><span data-stu-id="ed427-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ed427-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-219">Requirements</span></span>

|<span data-ttu-id="ed427-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-220">Requirement</span></span>| <span data-ttu-id="ed427-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-223">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-223">1.1</span></span>|
|[<span data-ttu-id="ed427-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="ed427-227">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="ed427-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="ed427-228">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="ed427-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-229">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-229">Type</span></span>

*   [<span data-ttu-id="ed427-230">HostType</span><span class="sxs-lookup"><span data-stu-id="ed427-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="ed427-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-231">Requirements</span></span>

|<span data-ttu-id="ed427-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-232">Requirement</span></span>| <span data-ttu-id="ed427-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-235">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-235">1.1</span></span>|
|[<span data-ttu-id="ed427-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="ed427-239">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="ed427-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="ed427-240">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="ed427-241">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="ed427-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="ed427-242">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications clientes Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="ed427-243">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-244">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-244">Type</span></span>

*   [<span data-ttu-id="ed427-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="ed427-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="ed427-246">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="ed427-246">Properties:</span></span>

|<span data-ttu-id="ed427-247">Nom</span><span class="sxs-lookup"><span data-stu-id="ed427-247">Name</span></span>| <span data-ttu-id="ed427-248">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-248">Type</span></span>| <span data-ttu-id="ed427-249">Description</span><span class="sxs-lookup"><span data-stu-id="ed427-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="ed427-250">String</span><span class="sxs-lookup"><span data-stu-id="ed427-250">String</span></span>|<span data-ttu-id="ed427-251">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="ed427-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="ed427-252">String</span><span class="sxs-lookup"><span data-stu-id="ed427-252">String</span></span>|<span data-ttu-id="ed427-253">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="ed427-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="ed427-254">String</span><span class="sxs-lookup"><span data-stu-id="ed427-254">String</span></span>|<span data-ttu-id="ed427-255">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="ed427-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="ed427-256">String</span><span class="sxs-lookup"><span data-stu-id="ed427-256">String</span></span>|<span data-ttu-id="ed427-257">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="ed427-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="ed427-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-258">Requirements</span></span>

|<span data-ttu-id="ed427-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-259">Requirement</span></span>| <span data-ttu-id="ed427-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-262">Aperçu</span><span class="sxs-lookup"><span data-stu-id="ed427-262">Preview</span></span>|
|[<span data-ttu-id="ed427-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-265">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="ed427-266">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="ed427-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="ed427-267">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="ed427-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-268">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-268">Type</span></span>

*   [<span data-ttu-id="ed427-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="ed427-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="ed427-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-270">Requirements</span></span>

|<span data-ttu-id="ed427-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-271">Requirement</span></span>| <span data-ttu-id="ed427-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-274">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-274">1.1</span></span>|
|[<span data-ttu-id="ed427-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-276">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-277">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="ed427-278">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="ed427-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="ed427-279">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="ed427-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-280">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-280">Type</span></span>

*   [<span data-ttu-id="ed427-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="ed427-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="ed427-282">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-282">Requirements</span></span>

|<span data-ttu-id="ed427-283">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-283">Requirement</span></span>| <span data-ttu-id="ed427-284">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-285">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-286">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-286">1.1</span></span>|
|[<span data-ttu-id="ed427-287">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-288">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ed427-289">Exemple</span><span class="sxs-lookup"><span data-stu-id="ed427-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="ed427-290">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="ed427-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="ed427-291">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="ed427-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="ed427-292">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="ed427-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-293">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-293">Type</span></span>

*   [<span data-ttu-id="ed427-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="ed427-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="ed427-295">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-295">Requirements</span></span>

|<span data-ttu-id="ed427-296">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-296">Requirement</span></span>| <span data-ttu-id="ed427-297">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-298">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-299">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-299">1.1</span></span>|
|[<span data-ttu-id="ed427-300">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="ed427-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="ed427-301">Restreinte</span><span class="sxs-lookup"><span data-stu-id="ed427-301">Restricted</span></span>|
|[<span data-ttu-id="ed427-302">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-303">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="ed427-304">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="ed427-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="ed427-305">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="ed427-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="ed427-306">Type</span><span class="sxs-lookup"><span data-stu-id="ed427-306">Type</span></span>

*   [<span data-ttu-id="ed427-307">UI</span><span class="sxs-lookup"><span data-stu-id="ed427-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="ed427-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="ed427-308">Requirements</span></span>

|<span data-ttu-id="ed427-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="ed427-309">Requirement</span></span>| <span data-ttu-id="ed427-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="ed427-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="ed427-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="ed427-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="ed427-312">1.1</span><span class="sxs-lookup"><span data-stu-id="ed427-312">1.1</span></span>|
|[<span data-ttu-id="ed427-313">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="ed427-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="ed427-314">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="ed427-314">Compose or Read</span></span>|
