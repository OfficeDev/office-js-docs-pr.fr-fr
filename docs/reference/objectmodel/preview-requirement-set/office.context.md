---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9c2c661ce870e2007bd891aee040c6b3564f7b9e
ms.sourcegitcommit: a0262ea40cd23f221e69bcb0223110f011265d13
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/17/2020
ms.locfileid: "42688902"
---
# <a name="context"></a><span data-ttu-id="af54e-102">context</span><span class="sxs-lookup"><span data-stu-id="af54e-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="af54e-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="af54e-103">[Office](office.md).context</span></span>

<span data-ttu-id="af54e-104">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="af54e-105">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="af54e-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="af54e-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-106">Requirements</span></span>

|<span data-ttu-id="af54e-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-107">Requirement</span></span>| <span data-ttu-id="af54e-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-110">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-110">1.1</span></span>|
|[<span data-ttu-id="af54e-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="af54e-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="af54e-113">Properties</span></span>

| <span data-ttu-id="af54e-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="af54e-114">Property</span></span> | <span data-ttu-id="af54e-115">Modes</span><span class="sxs-lookup"><span data-stu-id="af54e-115">Modes</span></span> | <span data-ttu-id="af54e-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="af54e-116">Return type</span></span> | <span data-ttu-id="af54e-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="af54e-117">Minimum</span></span><br><span data-ttu-id="af54e-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="af54e-119">auth</span><span class="sxs-lookup"><span data-stu-id="af54e-119">auth</span></span>](#auth-auth) | <span data-ttu-id="af54e-120">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-120">Compose</span></span><br><span data-ttu-id="af54e-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-121">Read</span></span> | [<span data-ttu-id="af54e-122">Auth</span><span class="sxs-lookup"><span data-stu-id="af54e-122">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="af54e-123">Aperçu</span><span class="sxs-lookup"><span data-stu-id="af54e-123">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="af54e-124">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="af54e-124">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="af54e-125">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-125">Compose</span></span><br><span data-ttu-id="af54e-126">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-126">Read</span></span> | <span data-ttu-id="af54e-127">Chaîne</span><span class="sxs-lookup"><span data-stu-id="af54e-127">String</span></span> | [<span data-ttu-id="af54e-128">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-129">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="af54e-129">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="af54e-130">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-130">Compose</span></span><br><span data-ttu-id="af54e-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-131">Read</span></span> | [<span data-ttu-id="af54e-132">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="af54e-132">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="af54e-133">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-134">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="af54e-134">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="af54e-135">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-135">Compose</span></span><br><span data-ttu-id="af54e-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-136">Read</span></span> | <span data-ttu-id="af54e-137">Chaîne</span><span class="sxs-lookup"><span data-stu-id="af54e-137">String</span></span> | [<span data-ttu-id="af54e-138">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-139">hote</span><span class="sxs-lookup"><span data-stu-id="af54e-139">host</span></span>](#host-hosttype) | <span data-ttu-id="af54e-140">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-140">Compose</span></span><br><span data-ttu-id="af54e-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-141">Read</span></span> | [<span data-ttu-id="af54e-142">HostType</span><span class="sxs-lookup"><span data-stu-id="af54e-142">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="af54e-143">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-144">mailbox</span><span class="sxs-lookup"><span data-stu-id="af54e-144">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="af54e-145">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-145">Compose</span></span><br><span data-ttu-id="af54e-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-146">Read</span></span> | [<span data-ttu-id="af54e-147">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-147">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="af54e-148">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-149">officeTheme</span><span class="sxs-lookup"><span data-stu-id="af54e-149">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="af54e-150">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-150">Compose</span></span><br><span data-ttu-id="af54e-151">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-151">Read</span></span> | [<span data-ttu-id="af54e-152">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="af54e-152">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="af54e-153">Aperçu</span><span class="sxs-lookup"><span data-stu-id="af54e-153">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="af54e-154">plateforme</span><span class="sxs-lookup"><span data-stu-id="af54e-154">platform</span></span>](#platform-platformtype) | <span data-ttu-id="af54e-155">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-155">Compose</span></span><br><span data-ttu-id="af54e-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-156">Read</span></span> | [<span data-ttu-id="af54e-157">PlatformType</span><span class="sxs-lookup"><span data-stu-id="af54e-157">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="af54e-158">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-159">requise</span><span class="sxs-lookup"><span data-stu-id="af54e-159">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="af54e-160">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-160">Compose</span></span><br><span data-ttu-id="af54e-161">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-161">Read</span></span> | [<span data-ttu-id="af54e-162">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="af54e-162">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="af54e-163">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-164">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="af54e-164">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="af54e-165">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-165">Compose</span></span><br><span data-ttu-id="af54e-166">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-166">Read</span></span> | [<span data-ttu-id="af54e-167">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="af54e-167">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="af54e-168">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-168">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="af54e-169">ui</span><span class="sxs-lookup"><span data-stu-id="af54e-169">ui</span></span>](#ui-ui) | <span data-ttu-id="af54e-170">Composition</span><span class="sxs-lookup"><span data-stu-id="af54e-170">Compose</span></span><br><span data-ttu-id="af54e-171">Lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-171">Read</span></span> | [<span data-ttu-id="af54e-172">UI</span><span class="sxs-lookup"><span data-stu-id="af54e-172">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="af54e-173">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-173">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="af54e-174">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="af54e-174">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="af54e-175">AUTH : [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="af54e-175">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="af54e-176">Prend en charge l’authentification [unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’hôte Office d’obtenir un jeton d’accès à l’application Web du complément.</span><span class="sxs-lookup"><span data-stu-id="af54e-176">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="af54e-177">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="af54e-177">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-178">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-178">Type</span></span>

*   [<span data-ttu-id="af54e-179">Auth</span><span class="sxs-lookup"><span data-stu-id="af54e-179">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="af54e-180">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-180">Requirements</span></span>

|<span data-ttu-id="af54e-181">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-181">Requirement</span></span>| <span data-ttu-id="af54e-182">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-183">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-183">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-184">Aperçu</span><span class="sxs-lookup"><span data-stu-id="af54e-184">Preview</span></span>|
|[<span data-ttu-id="af54e-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-185">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-186">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-186">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-187">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-187">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="af54e-188">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="af54e-188">contentLanguage: String</span></span>

<span data-ttu-id="af54e-189">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="af54e-189">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="af54e-190">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-190">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-191">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-191">Type</span></span>

*   <span data-ttu-id="af54e-192">String</span><span class="sxs-lookup"><span data-stu-id="af54e-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af54e-193">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-193">Requirements</span></span>

|<span data-ttu-id="af54e-194">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-194">Requirement</span></span>| <span data-ttu-id="af54e-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-196">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-196">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-197">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-197">1.1</span></span>|
|[<span data-ttu-id="af54e-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-198">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-199">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-199">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-200">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-200">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="af54e-201">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="af54e-201">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="af54e-202">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="af54e-202">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-203">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-203">Type</span></span>

*   [<span data-ttu-id="af54e-204">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="af54e-204">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="af54e-205">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-205">Requirements</span></span>

|<span data-ttu-id="af54e-206">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-206">Requirement</span></span>| <span data-ttu-id="af54e-207">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-208">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-208">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-209">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-209">1.1</span></span>|
|[<span data-ttu-id="af54e-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-210">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-211">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-211">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-212">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-212">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="af54e-213">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="af54e-213">displayLanguage: String</span></span>

<span data-ttu-id="af54e-214">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-214">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="af54e-215">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-216">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-216">Type</span></span>

*   <span data-ttu-id="af54e-217">String</span><span class="sxs-lookup"><span data-stu-id="af54e-217">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="af54e-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-218">Requirements</span></span>

|<span data-ttu-id="af54e-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-219">Requirement</span></span>| <span data-ttu-id="af54e-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-222">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-222">1.1</span></span>|
|[<span data-ttu-id="af54e-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-224">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-225">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="af54e-226">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="af54e-226">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="af54e-227">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="af54e-227">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-228">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-228">Type</span></span>

*   [<span data-ttu-id="af54e-229">HostType</span><span class="sxs-lookup"><span data-stu-id="af54e-229">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="af54e-230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-230">Requirements</span></span>

|<span data-ttu-id="af54e-231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-231">Requirement</span></span>| <span data-ttu-id="af54e-232">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-234">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-234">1.1</span></span>|
|[<span data-ttu-id="af54e-235">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-236">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-236">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-237">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-237">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="af54e-238">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="af54e-238">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="af54e-239">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-239">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="af54e-240">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="af54e-240">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="af54e-241">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications hôtes Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-241">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="af54e-242">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-243">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-243">Type</span></span>

*   [<span data-ttu-id="af54e-244">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="af54e-244">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="af54e-245">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="af54e-245">Properties:</span></span>

|<span data-ttu-id="af54e-246">Nom</span><span class="sxs-lookup"><span data-stu-id="af54e-246">Name</span></span>| <span data-ttu-id="af54e-247">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-247">Type</span></span>| <span data-ttu-id="af54e-248">Description</span><span class="sxs-lookup"><span data-stu-id="af54e-248">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="af54e-249">Chaîne</span><span class="sxs-lookup"><span data-stu-id="af54e-249">String</span></span>|<span data-ttu-id="af54e-250">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="af54e-250">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="af54e-251">Chaîne</span><span class="sxs-lookup"><span data-stu-id="af54e-251">String</span></span>|<span data-ttu-id="af54e-252">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="af54e-252">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="af54e-253">String</span><span class="sxs-lookup"><span data-stu-id="af54e-253">String</span></span>|<span data-ttu-id="af54e-254">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="af54e-254">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="af54e-255">String</span><span class="sxs-lookup"><span data-stu-id="af54e-255">String</span></span>|<span data-ttu-id="af54e-256">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="af54e-256">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="af54e-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-257">Requirements</span></span>

|<span data-ttu-id="af54e-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-258">Requirement</span></span>| <span data-ttu-id="af54e-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-261">Aperçu</span><span class="sxs-lookup"><span data-stu-id="af54e-261">Preview</span></span>|
|[<span data-ttu-id="af54e-262">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-263">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-263">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-264">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-264">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="af54e-265">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="af54e-265">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="af54e-266">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="af54e-266">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-267">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-267">Type</span></span>

*   [<span data-ttu-id="af54e-268">PlatformType</span><span class="sxs-lookup"><span data-stu-id="af54e-268">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="af54e-269">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-269">Requirements</span></span>

|<span data-ttu-id="af54e-270">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-270">Requirement</span></span>| <span data-ttu-id="af54e-271">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-272">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-272">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-273">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-273">1.1</span></span>|
|[<span data-ttu-id="af54e-274">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-274">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-275">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-275">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-276">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-276">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="af54e-277">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="af54e-277">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="af54e-278">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="af54e-278">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-279">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-279">Type</span></span>

*   [<span data-ttu-id="af54e-280">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="af54e-280">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="af54e-281">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-281">Requirements</span></span>

|<span data-ttu-id="af54e-282">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-282">Requirement</span></span>| <span data-ttu-id="af54e-283">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-284">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-284">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-285">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-285">1.1</span></span>|
|[<span data-ttu-id="af54e-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-286">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-287">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="af54e-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="af54e-288">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="af54e-289">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="af54e-289">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="af54e-290">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="af54e-290">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="af54e-291">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="af54e-291">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-292">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-292">Type</span></span>

*   [<span data-ttu-id="af54e-293">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="af54e-293">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="af54e-294">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-294">Requirements</span></span>

|<span data-ttu-id="af54e-295">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-295">Requirement</span></span>| <span data-ttu-id="af54e-296">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-297">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-297">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-298">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-298">1.1</span></span>|
|[<span data-ttu-id="af54e-299">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="af54e-299">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="af54e-300">Restreinte</span><span class="sxs-lookup"><span data-stu-id="af54e-300">Restricted</span></span>|
|[<span data-ttu-id="af54e-301">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-301">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-302">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-302">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="af54e-303">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="af54e-303">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="af54e-304">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="af54e-304">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="af54e-305">Type</span><span class="sxs-lookup"><span data-stu-id="af54e-305">Type</span></span>

*   [<span data-ttu-id="af54e-306">UI</span><span class="sxs-lookup"><span data-stu-id="af54e-306">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="af54e-307">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="af54e-307">Requirements</span></span>

|<span data-ttu-id="af54e-308">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="af54e-308">Requirement</span></span>| <span data-ttu-id="af54e-309">Valeur</span><span class="sxs-lookup"><span data-stu-id="af54e-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="af54e-310">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="af54e-310">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="af54e-311">1.1</span><span class="sxs-lookup"><span data-stu-id="af54e-311">1.1</span></span>|
|[<span data-ttu-id="af54e-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="af54e-312">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="af54e-313">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="af54e-313">Compose or Read</span></span>|
