---
title: Ensemble de conditions requises pour Office. Context-preview
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook à l’aide de l’ensemble de conditions requises pour l’API de boîte aux lettres.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 8370df907aa3ab0534254057860c187cec583e6c
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570785"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="b52d3-103">contexte (ensemble de conditions requises pour la boîte aux lettres)</span><span class="sxs-lookup"><span data-stu-id="b52d3-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="b52d3-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b52d3-104">[Office](office.md).context</span></span>

<span data-ttu-id="b52d3-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b52d3-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="b52d3-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b52d3-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-107">Requirements</span></span>

|<span data-ttu-id="b52d3-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-108">Requirement</span></span>| <span data-ttu-id="b52d3-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-111">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-111">1.1</span></span>|
|[<span data-ttu-id="b52d3-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="b52d3-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="b52d3-114">Properties</span></span>

| <span data-ttu-id="b52d3-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="b52d3-115">Property</span></span> | <span data-ttu-id="b52d3-116">Modes</span><span class="sxs-lookup"><span data-stu-id="b52d3-116">Modes</span></span> | <span data-ttu-id="b52d3-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="b52d3-117">Return type</span></span> | <span data-ttu-id="b52d3-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="b52d3-118">Minimum</span></span><br><span data-ttu-id="b52d3-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b52d3-120">auth</span><span class="sxs-lookup"><span data-stu-id="b52d3-120">auth</span></span>](#auth-auth) | <span data-ttu-id="b52d3-121">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-121">Compose</span></span><br><span data-ttu-id="b52d3-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-122">Read</span></span> | [<span data-ttu-id="b52d3-123">Auth</span><span class="sxs-lookup"><span data-stu-id="b52d3-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-124">Ensembles 1,3</span><span class="sxs-lookup"><span data-stu-id="b52d3-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="b52d3-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b52d3-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b52d3-126">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-126">Compose</span></span><br><span data-ttu-id="b52d3-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-127">Read</span></span> | <span data-ttu-id="b52d3-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b52d3-128">String</span></span> | [<span data-ttu-id="b52d3-129">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b52d3-130">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="b52d3-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b52d3-131">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-131">Compose</span></span><br><span data-ttu-id="b52d3-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-132">Read</span></span> | [<span data-ttu-id="b52d3-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b52d3-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b52d3-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b52d3-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b52d3-136">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-136">Compose</span></span><br><span data-ttu-id="b52d3-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-137">Read</span></span> | <span data-ttu-id="b52d3-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b52d3-138">String</span></span> | [<span data-ttu-id="b52d3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b52d3-140">hote</span><span class="sxs-lookup"><span data-stu-id="b52d3-140">host</span></span>](#host-hosttype) | <span data-ttu-id="b52d3-141">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-141">Compose</span></span><br><span data-ttu-id="b52d3-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-142">Read</span></span> | [<span data-ttu-id="b52d3-143">HostType</span><span class="sxs-lookup"><span data-stu-id="b52d3-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-144">1,5</span><span class="sxs-lookup"><span data-stu-id="b52d3-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b52d3-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="b52d3-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b52d3-146">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-146">Compose</span></span><br><span data-ttu-id="b52d3-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-147">Read</span></span> | [<span data-ttu-id="b52d3-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b52d3-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="b52d3-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="b52d3-151">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-151">Compose</span></span><br><span data-ttu-id="b52d3-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-152">Read</span></span> | [<span data-ttu-id="b52d3-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="b52d3-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-154">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b52d3-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="b52d3-155">plateforme</span><span class="sxs-lookup"><span data-stu-id="b52d3-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b52d3-156">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-156">Compose</span></span><br><span data-ttu-id="b52d3-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-157">Read</span></span> | [<span data-ttu-id="b52d3-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b52d3-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-159">1,5</span><span class="sxs-lookup"><span data-stu-id="b52d3-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b52d3-160">requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b52d3-161">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-161">Compose</span></span><br><span data-ttu-id="b52d3-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-162">Read</span></span> | [<span data-ttu-id="b52d3-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b52d3-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b52d3-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b52d3-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b52d3-166">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-166">Compose</span></span><br><span data-ttu-id="b52d3-167">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-167">Read</span></span> | [<span data-ttu-id="b52d3-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b52d3-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-169">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b52d3-170">ui</span><span class="sxs-lookup"><span data-stu-id="b52d3-170">ui</span></span>](#ui-ui) | <span data-ttu-id="b52d3-171">Composition</span><span class="sxs-lookup"><span data-stu-id="b52d3-171">Compose</span></span><br><span data-ttu-id="b52d3-172">Lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-172">Read</span></span> | [<span data-ttu-id="b52d3-173">UI</span><span class="sxs-lookup"><span data-stu-id="b52d3-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="b52d3-174">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b52d3-175">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="b52d3-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="b52d3-176">AUTH : [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="b52d3-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="b52d3-177">Prend en charge l’authentification [unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application Web du complément.</span><span class="sxs-lookup"><span data-stu-id="b52d3-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="b52d3-178">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="b52d3-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-179">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-179">Type</span></span>

*   [<span data-ttu-id="b52d3-180">Auth</span><span class="sxs-lookup"><span data-stu-id="b52d3-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="b52d3-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-181">Requirements</span></span>

|<span data-ttu-id="b52d3-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-182">Requirement</span></span>| <span data-ttu-id="b52d3-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-185">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b52d3-185">Preview</span></span>|
|[<span data-ttu-id="b52d3-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="b52d3-189">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="b52d3-189">contentLanguage: String</span></span>

<span data-ttu-id="b52d3-190">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="b52d3-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b52d3-191">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-192">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-192">Type</span></span>

*   <span data-ttu-id="b52d3-193">String</span><span class="sxs-lookup"><span data-stu-id="b52d3-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b52d3-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-194">Requirements</span></span>

|<span data-ttu-id="b52d3-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-195">Requirement</span></span>| <span data-ttu-id="b52d3-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-198">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-198">1.1</span></span>|
|[<span data-ttu-id="b52d3-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-201">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b52d3-202">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b52d3-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b52d3-203">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="b52d3-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-204">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-204">Type</span></span>

*   [<span data-ttu-id="b52d3-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b52d3-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b52d3-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-206">Requirements</span></span>

|<span data-ttu-id="b52d3-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-207">Requirement</span></span>| <span data-ttu-id="b52d3-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-210">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-210">1.1</span></span>|
|[<span data-ttu-id="b52d3-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-212">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b52d3-214">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="b52d3-214">displayLanguage: String</span></span>

<span data-ttu-id="b52d3-215">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="b52d3-216">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-217">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-217">Type</span></span>

*   <span data-ttu-id="b52d3-218">String</span><span class="sxs-lookup"><span data-stu-id="b52d3-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b52d3-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-219">Requirements</span></span>

|<span data-ttu-id="b52d3-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-220">Requirement</span></span>| <span data-ttu-id="b52d3-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-223">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-223">1.1</span></span>|
|[<span data-ttu-id="b52d3-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="b52d3-227">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b52d3-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b52d3-228">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="b52d3-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b52d3-229">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.</span><span class="sxs-lookup"><span data-stu-id="b52d3-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-230">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-230">Type</span></span>

*   [<span data-ttu-id="b52d3-231">HostType</span><span class="sxs-lookup"><span data-stu-id="b52d3-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b52d3-232">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-232">Requirements</span></span>

|<span data-ttu-id="b52d3-233">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-233">Requirement</span></span>| <span data-ttu-id="b52d3-234">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-235">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-236">1,5</span><span class="sxs-lookup"><span data-stu-id="b52d3-236">1.5</span></span>|
|[<span data-ttu-id="b52d3-237">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-238">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-239">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="b52d3-240">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="b52d3-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="b52d3-241">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="b52d3-242">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="b52d3-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="b52d3-243">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème** Office, qui est appliquée à toutes les applications clientes Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="b52d3-244">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-245">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-245">Type</span></span>

*   [<span data-ttu-id="b52d3-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="b52d3-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="b52d3-247">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="b52d3-247">Properties:</span></span>

|<span data-ttu-id="b52d3-248">Nom</span><span class="sxs-lookup"><span data-stu-id="b52d3-248">Name</span></span>| <span data-ttu-id="b52d3-249">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-249">Type</span></span>| <span data-ttu-id="b52d3-250">Description</span><span class="sxs-lookup"><span data-stu-id="b52d3-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="b52d3-251">Chaîne</span><span class="sxs-lookup"><span data-stu-id="b52d3-251">String</span></span>|<span data-ttu-id="b52d3-252">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="b52d3-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="b52d3-253">String</span><span class="sxs-lookup"><span data-stu-id="b52d3-253">String</span></span>|<span data-ttu-id="b52d3-254">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="b52d3-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="b52d3-255">String</span><span class="sxs-lookup"><span data-stu-id="b52d3-255">String</span></span>|<span data-ttu-id="b52d3-256">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="b52d3-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="b52d3-257">String</span><span class="sxs-lookup"><span data-stu-id="b52d3-257">String</span></span>|<span data-ttu-id="b52d3-258">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="b52d3-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b52d3-259">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-259">Requirements</span></span>

|<span data-ttu-id="b52d3-260">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-260">Requirement</span></span>| <span data-ttu-id="b52d3-261">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-262">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-263">Aperçu</span><span class="sxs-lookup"><span data-stu-id="b52d3-263">Preview</span></span>|
|[<span data-ttu-id="b52d3-264">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-265">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-266">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-266">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="b52d3-267">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b52d3-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b52d3-268">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="b52d3-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="b52d3-269">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="b52d3-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-270">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-270">Type</span></span>

*   [<span data-ttu-id="b52d3-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b52d3-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b52d3-272">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-272">Requirements</span></span>

|<span data-ttu-id="b52d3-273">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-273">Requirement</span></span>| <span data-ttu-id="b52d3-274">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-275">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-276">1,5</span><span class="sxs-lookup"><span data-stu-id="b52d3-276">1.5</span></span>|
|[<span data-ttu-id="b52d3-277">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-278">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-279">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b52d3-280">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b52d3-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b52d3-281">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="b52d3-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-282">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-282">Type</span></span>

*   [<span data-ttu-id="b52d3-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b52d3-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b52d3-284">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-284">Requirements</span></span>

|<span data-ttu-id="b52d3-285">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-285">Requirement</span></span>| <span data-ttu-id="b52d3-286">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-287">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-288">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-288">1.1</span></span>|
|[<span data-ttu-id="b52d3-289">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-290">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b52d3-291">Exemple</span><span class="sxs-lookup"><span data-stu-id="b52d3-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b52d3-292">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b52d3-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b52d3-293">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b52d3-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b52d3-294">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="b52d3-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-295">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-295">Type</span></span>

*   [<span data-ttu-id="b52d3-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b52d3-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b52d3-297">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-297">Requirements</span></span>

|<span data-ttu-id="b52d3-298">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-298">Requirement</span></span>| <span data-ttu-id="b52d3-299">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-300">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-301">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-301">1.1</span></span>|
|[<span data-ttu-id="b52d3-302">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b52d3-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b52d3-303">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b52d3-303">Restricted</span></span>|
|[<span data-ttu-id="b52d3-304">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-305">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b52d3-306">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b52d3-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b52d3-307">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="b52d3-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b52d3-308">Type</span><span class="sxs-lookup"><span data-stu-id="b52d3-308">Type</span></span>

*   [<span data-ttu-id="b52d3-309">UI</span><span class="sxs-lookup"><span data-stu-id="b52d3-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b52d3-310">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b52d3-310">Requirements</span></span>

|<span data-ttu-id="b52d3-311">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b52d3-311">Requirement</span></span>| <span data-ttu-id="b52d3-312">Valeur</span><span class="sxs-lookup"><span data-stu-id="b52d3-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="b52d3-313">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b52d3-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b52d3-314">1.1</span><span class="sxs-lookup"><span data-stu-id="b52d3-314">1.1</span></span>|
|[<span data-ttu-id="b52d3-315">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b52d3-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b52d3-316">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b52d3-316">Compose or Read</span></span>|
