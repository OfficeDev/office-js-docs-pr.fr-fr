---
title: Ensemble de conditions requises pour Office. Context-preview
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook à l’aide de l’ensemble de conditions requises pour l’API de boîte aux lettres.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: c61769cb1ae98097ffabb8b3ef19b2f82257c2b1
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890864"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="32977-103">contexte (ensemble de conditions requises pour la boîte aux lettres)</span><span class="sxs-lookup"><span data-stu-id="32977-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="32977-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="32977-104">[Office](office.md).context</span></span>

<span data-ttu-id="32977-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="32977-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="32977-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="32977-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="32977-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-107">Requirements</span></span>

|<span data-ttu-id="32977-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-108">Requirement</span></span>| <span data-ttu-id="32977-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-111">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-111">1.1</span></span>|
|[<span data-ttu-id="32977-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="32977-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="32977-114">Properties</span></span>

| <span data-ttu-id="32977-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="32977-115">Property</span></span> | <span data-ttu-id="32977-116">Modes</span><span class="sxs-lookup"><span data-stu-id="32977-116">Modes</span></span> | <span data-ttu-id="32977-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="32977-117">Return type</span></span> | <span data-ttu-id="32977-118">Minimale</span><span class="sxs-lookup"><span data-stu-id="32977-118">Minimum</span></span><br><span data-ttu-id="32977-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="32977-120">auth</span><span class="sxs-lookup"><span data-stu-id="32977-120">auth</span></span>](#auth-auth) | <span data-ttu-id="32977-121">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-121">Compose</span></span><br><span data-ttu-id="32977-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-122">Read</span></span> | [<span data-ttu-id="32977-123">Auth</span><span class="sxs-lookup"><span data-stu-id="32977-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="32977-124">Aperçu</span><span class="sxs-lookup"><span data-stu-id="32977-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="32977-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="32977-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="32977-126">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-126">Compose</span></span><br><span data-ttu-id="32977-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-127">Read</span></span> | <span data-ttu-id="32977-128">String</span><span class="sxs-lookup"><span data-stu-id="32977-128">String</span></span> | [<span data-ttu-id="32977-129">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-130">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="32977-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="32977-131">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-131">Compose</span></span><br><span data-ttu-id="32977-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-132">Read</span></span> | [<span data-ttu-id="32977-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="32977-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="32977-134">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="32977-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="32977-136">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-136">Compose</span></span><br><span data-ttu-id="32977-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-137">Read</span></span> | <span data-ttu-id="32977-138">String</span><span class="sxs-lookup"><span data-stu-id="32977-138">String</span></span> | [<span data-ttu-id="32977-139">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-140">hote</span><span class="sxs-lookup"><span data-stu-id="32977-140">host</span></span>](#host-hosttype) | <span data-ttu-id="32977-141">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-141">Compose</span></span><br><span data-ttu-id="32977-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-142">Read</span></span> | [<span data-ttu-id="32977-143">HostType</span><span class="sxs-lookup"><span data-stu-id="32977-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="32977-144">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="32977-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="32977-146">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-146">Compose</span></span><br><span data-ttu-id="32977-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-147">Read</span></span> | [<span data-ttu-id="32977-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="32977-149">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="32977-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="32977-151">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-151">Compose</span></span><br><span data-ttu-id="32977-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-152">Read</span></span> | [<span data-ttu-id="32977-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="32977-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="32977-154">Aperçu</span><span class="sxs-lookup"><span data-stu-id="32977-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="32977-155">plateforme</span><span class="sxs-lookup"><span data-stu-id="32977-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="32977-156">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-156">Compose</span></span><br><span data-ttu-id="32977-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-157">Read</span></span> | [<span data-ttu-id="32977-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="32977-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="32977-159">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-160">requise</span><span class="sxs-lookup"><span data-stu-id="32977-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="32977-161">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-161">Compose</span></span><br><span data-ttu-id="32977-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-162">Read</span></span> | [<span data-ttu-id="32977-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="32977-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="32977-164">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="32977-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="32977-166">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-166">Compose</span></span><br><span data-ttu-id="32977-167">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-167">Read</span></span> | [<span data-ttu-id="32977-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="32977-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="32977-169">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="32977-170">ui</span><span class="sxs-lookup"><span data-stu-id="32977-170">ui</span></span>](#ui-ui) | <span data-ttu-id="32977-171">Composition</span><span class="sxs-lookup"><span data-stu-id="32977-171">Compose</span></span><br><span data-ttu-id="32977-172">Lecture</span><span class="sxs-lookup"><span data-stu-id="32977-172">Read</span></span> | [<span data-ttu-id="32977-173">UI</span><span class="sxs-lookup"><span data-stu-id="32977-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="32977-174">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="32977-175">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="32977-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="32977-176">AUTH : [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="32977-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="32977-177">Prend en charge l’authentification [unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’hôte Office d’obtenir un jeton d’accès à l’application Web du complément.</span><span class="sxs-lookup"><span data-stu-id="32977-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="32977-178">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="32977-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-179">Type</span><span class="sxs-lookup"><span data-stu-id="32977-179">Type</span></span>

*   [<span data-ttu-id="32977-180">Auth</span><span class="sxs-lookup"><span data-stu-id="32977-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="32977-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-181">Requirements</span></span>

|<span data-ttu-id="32977-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-182">Requirement</span></span>| <span data-ttu-id="32977-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-185">Aperçu</span><span class="sxs-lookup"><span data-stu-id="32977-185">Preview</span></span>|
|[<span data-ttu-id="32977-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="32977-189">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="32977-189">contentLanguage: String</span></span>

<span data-ttu-id="32977-190">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="32977-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="32977-191">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="32977-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-192">Type</span><span class="sxs-lookup"><span data-stu-id="32977-192">Type</span></span>

*   <span data-ttu-id="32977-193">String</span><span class="sxs-lookup"><span data-stu-id="32977-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32977-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-194">Requirements</span></span>

|<span data-ttu-id="32977-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-195">Requirement</span></span>| <span data-ttu-id="32977-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-198">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-198">1.1</span></span>|
|[<span data-ttu-id="32977-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-201">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="32977-202">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="32977-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="32977-203">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="32977-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-204">Type</span><span class="sxs-lookup"><span data-stu-id="32977-204">Type</span></span>

*   [<span data-ttu-id="32977-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="32977-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="32977-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-206">Requirements</span></span>

|<span data-ttu-id="32977-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-207">Requirement</span></span>| <span data-ttu-id="32977-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-210">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-210">1.1</span></span>|
|[<span data-ttu-id="32977-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-212">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="32977-214">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="32977-214">displayLanguage: String</span></span>

<span data-ttu-id="32977-215">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="32977-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="32977-216">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="32977-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-217">Type</span><span class="sxs-lookup"><span data-stu-id="32977-217">Type</span></span>

*   <span data-ttu-id="32977-218">String</span><span class="sxs-lookup"><span data-stu-id="32977-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="32977-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-219">Requirements</span></span>

|<span data-ttu-id="32977-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-220">Requirement</span></span>| <span data-ttu-id="32977-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-223">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-223">1.1</span></span>|
|[<span data-ttu-id="32977-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="32977-227">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="32977-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="32977-228">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="32977-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-229">Type</span><span class="sxs-lookup"><span data-stu-id="32977-229">Type</span></span>

*   [<span data-ttu-id="32977-230">HostType</span><span class="sxs-lookup"><span data-stu-id="32977-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="32977-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-231">Requirements</span></span>

|<span data-ttu-id="32977-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-232">Requirement</span></span>| <span data-ttu-id="32977-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-235">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-235">1.1</span></span>|
|[<span data-ttu-id="32977-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="32977-239">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="32977-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="32977-240">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="32977-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="32977-241">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="32977-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="32977-242">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications hôtes Office.</span><span class="sxs-lookup"><span data-stu-id="32977-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="32977-243">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="32977-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-244">Type</span><span class="sxs-lookup"><span data-stu-id="32977-244">Type</span></span>

*   [<span data-ttu-id="32977-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="32977-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="32977-246">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="32977-246">Properties:</span></span>

|<span data-ttu-id="32977-247">Nom</span><span class="sxs-lookup"><span data-stu-id="32977-247">Name</span></span>| <span data-ttu-id="32977-248">Type</span><span class="sxs-lookup"><span data-stu-id="32977-248">Type</span></span>| <span data-ttu-id="32977-249">Description</span><span class="sxs-lookup"><span data-stu-id="32977-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="32977-250">String</span><span class="sxs-lookup"><span data-stu-id="32977-250">String</span></span>|<span data-ttu-id="32977-251">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="32977-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="32977-252">String</span><span class="sxs-lookup"><span data-stu-id="32977-252">String</span></span>|<span data-ttu-id="32977-253">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="32977-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="32977-254">String</span><span class="sxs-lookup"><span data-stu-id="32977-254">String</span></span>|<span data-ttu-id="32977-255">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="32977-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="32977-256">String</span><span class="sxs-lookup"><span data-stu-id="32977-256">String</span></span>|<span data-ttu-id="32977-257">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="32977-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="32977-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-258">Requirements</span></span>

|<span data-ttu-id="32977-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-259">Requirement</span></span>| <span data-ttu-id="32977-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-262">Aperçu</span><span class="sxs-lookup"><span data-stu-id="32977-262">Preview</span></span>|
|[<span data-ttu-id="32977-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-265">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="32977-266">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="32977-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="32977-267">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="32977-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-268">Type</span><span class="sxs-lookup"><span data-stu-id="32977-268">Type</span></span>

*   [<span data-ttu-id="32977-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="32977-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="32977-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-270">Requirements</span></span>

|<span data-ttu-id="32977-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-271">Requirement</span></span>| <span data-ttu-id="32977-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-274">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-274">1.1</span></span>|
|[<span data-ttu-id="32977-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-276">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-277">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="32977-278">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="32977-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="32977-279">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="32977-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-280">Type</span><span class="sxs-lookup"><span data-stu-id="32977-280">Type</span></span>

*   [<span data-ttu-id="32977-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="32977-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="32977-282">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-282">Requirements</span></span>

|<span data-ttu-id="32977-283">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-283">Requirement</span></span>| <span data-ttu-id="32977-284">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-285">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-286">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-286">1.1</span></span>|
|[<span data-ttu-id="32977-287">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-288">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="32977-289">Exemple</span><span class="sxs-lookup"><span data-stu-id="32977-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="32977-290">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="32977-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="32977-291">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="32977-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="32977-292">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="32977-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-293">Type</span><span class="sxs-lookup"><span data-stu-id="32977-293">Type</span></span>

*   [<span data-ttu-id="32977-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="32977-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="32977-295">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-295">Requirements</span></span>

|<span data-ttu-id="32977-296">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-296">Requirement</span></span>| <span data-ttu-id="32977-297">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-298">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-299">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-299">1.1</span></span>|
|[<span data-ttu-id="32977-300">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="32977-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="32977-301">Restreinte</span><span class="sxs-lookup"><span data-stu-id="32977-301">Restricted</span></span>|
|[<span data-ttu-id="32977-302">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-303">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="32977-304">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="32977-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="32977-305">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="32977-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="32977-306">Type</span><span class="sxs-lookup"><span data-stu-id="32977-306">Type</span></span>

*   [<span data-ttu-id="32977-307">UI</span><span class="sxs-lookup"><span data-stu-id="32977-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="32977-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="32977-308">Requirements</span></span>

|<span data-ttu-id="32977-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="32977-309">Requirement</span></span>| <span data-ttu-id="32977-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="32977-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="32977-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="32977-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="32977-312">1.1</span><span class="sxs-lookup"><span data-stu-id="32977-312">1.1</span></span>|
|[<span data-ttu-id="32977-313">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="32977-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="32977-314">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="32977-314">Compose or Read</span></span>|
