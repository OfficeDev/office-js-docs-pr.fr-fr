---
title: Ensemble de conditions requises pour Office. Context-preview
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook à l’aide de l’ensemble de conditions requises pour l’API de boîte aux lettres.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 5987f81b0b4790b74bde092fc3de44df4fa3ed16
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293813"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="47b77-103">contexte (ensemble de conditions requises pour la boîte aux lettres)</span><span class="sxs-lookup"><span data-stu-id="47b77-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="47b77-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="47b77-104">[Office](office.md).context</span></span>

<span data-ttu-id="47b77-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="47b77-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="47b77-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="47b77-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-107">Requirements</span></span>

|<span data-ttu-id="47b77-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-108">Requirement</span></span>| <span data-ttu-id="47b77-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-111">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-111">1.1</span></span>|
|[<span data-ttu-id="47b77-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="47b77-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="47b77-114">Properties</span></span>

| <span data-ttu-id="47b77-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="47b77-115">Property</span></span> | <span data-ttu-id="47b77-116">Modes</span><span class="sxs-lookup"><span data-stu-id="47b77-116">Modes</span></span> | <span data-ttu-id="47b77-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="47b77-117">Return type</span></span> | <span data-ttu-id="47b77-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="47b77-118">Minimum</span></span><br><span data-ttu-id="47b77-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="47b77-120">auth</span><span class="sxs-lookup"><span data-stu-id="47b77-120">auth</span></span>](#auth-auth) | <span data-ttu-id="47b77-121">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-121">Compose</span></span><br><span data-ttu-id="47b77-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-122">Read</span></span> | [<span data-ttu-id="47b77-123">Auth</span><span class="sxs-lookup"><span data-stu-id="47b77-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="47b77-124">Aperçu</span><span class="sxs-lookup"><span data-stu-id="47b77-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="47b77-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="47b77-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="47b77-126">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-126">Compose</span></span><br><span data-ttu-id="47b77-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-127">Read</span></span> | <span data-ttu-id="47b77-128">String</span><span class="sxs-lookup"><span data-stu-id="47b77-128">String</span></span> | [<span data-ttu-id="47b77-129">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-130">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="47b77-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="47b77-131">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-131">Compose</span></span><br><span data-ttu-id="47b77-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-132">Read</span></span> | [<span data-ttu-id="47b77-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="47b77-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="47b77-134">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="47b77-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="47b77-136">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-136">Compose</span></span><br><span data-ttu-id="47b77-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-137">Read</span></span> | <span data-ttu-id="47b77-138">String</span><span class="sxs-lookup"><span data-stu-id="47b77-138">String</span></span> | [<span data-ttu-id="47b77-139">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-140">hote</span><span class="sxs-lookup"><span data-stu-id="47b77-140">host</span></span>](#host-hosttype) | <span data-ttu-id="47b77-141">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-141">Compose</span></span><br><span data-ttu-id="47b77-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-142">Read</span></span> | [<span data-ttu-id="47b77-143">HostType</span><span class="sxs-lookup"><span data-stu-id="47b77-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="47b77-144">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="47b77-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="47b77-146">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-146">Compose</span></span><br><span data-ttu-id="47b77-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-147">Read</span></span> | [<span data-ttu-id="47b77-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="47b77-149">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="47b77-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="47b77-151">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-151">Compose</span></span><br><span data-ttu-id="47b77-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-152">Read</span></span> | [<span data-ttu-id="47b77-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="47b77-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="47b77-154">Aperçu</span><span class="sxs-lookup"><span data-stu-id="47b77-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="47b77-155">platform</span><span class="sxs-lookup"><span data-stu-id="47b77-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="47b77-156">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-156">Compose</span></span><br><span data-ttu-id="47b77-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-157">Read</span></span> | [<span data-ttu-id="47b77-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="47b77-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="47b77-159">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-160">requise</span><span class="sxs-lookup"><span data-stu-id="47b77-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="47b77-161">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-161">Compose</span></span><br><span data-ttu-id="47b77-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-162">Read</span></span> | [<span data-ttu-id="47b77-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="47b77-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="47b77-164">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="47b77-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="47b77-166">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-166">Compose</span></span><br><span data-ttu-id="47b77-167">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-167">Read</span></span> | [<span data-ttu-id="47b77-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="47b77-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="47b77-169">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="47b77-170">ui</span><span class="sxs-lookup"><span data-stu-id="47b77-170">ui</span></span>](#ui-ui) | <span data-ttu-id="47b77-171">Composition</span><span class="sxs-lookup"><span data-stu-id="47b77-171">Compose</span></span><br><span data-ttu-id="47b77-172">Lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-172">Read</span></span> | [<span data-ttu-id="47b77-173">UI</span><span class="sxs-lookup"><span data-stu-id="47b77-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="47b77-174">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="47b77-175">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="47b77-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="47b77-176">AUTH : [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="47b77-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="47b77-177">Prend en charge l’authentification [unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application Web du complément.</span><span class="sxs-lookup"><span data-stu-id="47b77-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="47b77-178">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="47b77-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-179">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-179">Type</span></span>

*   [<span data-ttu-id="47b77-180">Auth</span><span class="sxs-lookup"><span data-stu-id="47b77-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="47b77-181">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-181">Requirements</span></span>

|<span data-ttu-id="47b77-182">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-182">Requirement</span></span>| <span data-ttu-id="47b77-183">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-184">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-185">Aperçu</span><span class="sxs-lookup"><span data-stu-id="47b77-185">Preview</span></span>|
|[<span data-ttu-id="47b77-186">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-187">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-188">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="47b77-189">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="47b77-189">contentLanguage: String</span></span>

<span data-ttu-id="47b77-190">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="47b77-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="47b77-191">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-192">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-192">Type</span></span>

*   <span data-ttu-id="47b77-193">String</span><span class="sxs-lookup"><span data-stu-id="47b77-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="47b77-194">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-194">Requirements</span></span>

|<span data-ttu-id="47b77-195">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-195">Requirement</span></span>| <span data-ttu-id="47b77-196">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-197">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-198">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-198">1.1</span></span>|
|[<span data-ttu-id="47b77-199">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-200">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-201">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="47b77-202">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="47b77-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="47b77-203">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="47b77-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-204">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-204">Type</span></span>

*   [<span data-ttu-id="47b77-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="47b77-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="47b77-206">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-206">Requirements</span></span>

|<span data-ttu-id="47b77-207">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-207">Requirement</span></span>| <span data-ttu-id="47b77-208">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-209">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-210">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-210">1.1</span></span>|
|[<span data-ttu-id="47b77-211">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-212">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-213">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="47b77-214">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="47b77-214">displayLanguage: String</span></span>

<span data-ttu-id="47b77-215">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="47b77-216">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-217">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-217">Type</span></span>

*   <span data-ttu-id="47b77-218">String</span><span class="sxs-lookup"><span data-stu-id="47b77-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="47b77-219">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-219">Requirements</span></span>

|<span data-ttu-id="47b77-220">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-220">Requirement</span></span>| <span data-ttu-id="47b77-221">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-222">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-223">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-223">1.1</span></span>|
|[<span data-ttu-id="47b77-224">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-225">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-226">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="47b77-227">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="47b77-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="47b77-228">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="47b77-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-229">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-229">Type</span></span>

*   [<span data-ttu-id="47b77-230">HostType</span><span class="sxs-lookup"><span data-stu-id="47b77-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="47b77-231">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-231">Requirements</span></span>

|<span data-ttu-id="47b77-232">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-232">Requirement</span></span>| <span data-ttu-id="47b77-233">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-234">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-235">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-235">1.1</span></span>|
|[<span data-ttu-id="47b77-236">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-237">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-238">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="47b77-239">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="47b77-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="47b77-240">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="47b77-241">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="47b77-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="47b77-242">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications clientes Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="47b77-243">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-244">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-244">Type</span></span>

*   [<span data-ttu-id="47b77-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="47b77-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="47b77-246">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="47b77-246">Properties:</span></span>

|<span data-ttu-id="47b77-247">Nom</span><span class="sxs-lookup"><span data-stu-id="47b77-247">Name</span></span>| <span data-ttu-id="47b77-248">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-248">Type</span></span>| <span data-ttu-id="47b77-249">Description</span><span class="sxs-lookup"><span data-stu-id="47b77-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="47b77-250">String</span><span class="sxs-lookup"><span data-stu-id="47b77-250">String</span></span>|<span data-ttu-id="47b77-251">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="47b77-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="47b77-252">String</span><span class="sxs-lookup"><span data-stu-id="47b77-252">String</span></span>|<span data-ttu-id="47b77-253">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="47b77-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="47b77-254">String</span><span class="sxs-lookup"><span data-stu-id="47b77-254">String</span></span>|<span data-ttu-id="47b77-255">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="47b77-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="47b77-256">String</span><span class="sxs-lookup"><span data-stu-id="47b77-256">String</span></span>|<span data-ttu-id="47b77-257">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="47b77-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="47b77-258">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-258">Requirements</span></span>

|<span data-ttu-id="47b77-259">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-259">Requirement</span></span>| <span data-ttu-id="47b77-260">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-261">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-262">Aperçu</span><span class="sxs-lookup"><span data-stu-id="47b77-262">Preview</span></span>|
|[<span data-ttu-id="47b77-263">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-264">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-265">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="47b77-266">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="47b77-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="47b77-267">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="47b77-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-268">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-268">Type</span></span>

*   [<span data-ttu-id="47b77-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="47b77-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="47b77-270">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-270">Requirements</span></span>

|<span data-ttu-id="47b77-271">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-271">Requirement</span></span>| <span data-ttu-id="47b77-272">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-273">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-274">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-274">1.1</span></span>|
|[<span data-ttu-id="47b77-275">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-276">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-277">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="47b77-278">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="47b77-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="47b77-279">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="47b77-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-280">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-280">Type</span></span>

*   [<span data-ttu-id="47b77-281">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="47b77-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="47b77-282">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-282">Requirements</span></span>

|<span data-ttu-id="47b77-283">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-283">Requirement</span></span>| <span data-ttu-id="47b77-284">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-285">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-286">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-286">1.1</span></span>|
|[<span data-ttu-id="47b77-287">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-288">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="47b77-289">Exemple</span><span class="sxs-lookup"><span data-stu-id="47b77-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="47b77-290">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="47b77-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="47b77-291">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="47b77-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="47b77-292">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="47b77-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-293">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-293">Type</span></span>

*   [<span data-ttu-id="47b77-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="47b77-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="47b77-295">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-295">Requirements</span></span>

|<span data-ttu-id="47b77-296">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-296">Requirement</span></span>| <span data-ttu-id="47b77-297">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-298">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-299">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-299">1.1</span></span>|
|[<span data-ttu-id="47b77-300">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="47b77-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="47b77-301">Restreinte</span><span class="sxs-lookup"><span data-stu-id="47b77-301">Restricted</span></span>|
|[<span data-ttu-id="47b77-302">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-303">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="47b77-304">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="47b77-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="47b77-305">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="47b77-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="47b77-306">Type</span><span class="sxs-lookup"><span data-stu-id="47b77-306">Type</span></span>

*   [<span data-ttu-id="47b77-307">UI</span><span class="sxs-lookup"><span data-stu-id="47b77-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="47b77-308">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="47b77-308">Requirements</span></span>

|<span data-ttu-id="47b77-309">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="47b77-309">Requirement</span></span>| <span data-ttu-id="47b77-310">Valeur</span><span class="sxs-lookup"><span data-stu-id="47b77-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="47b77-311">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="47b77-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="47b77-312">1.1</span><span class="sxs-lookup"><span data-stu-id="47b77-312">1.1</span></span>|
|[<span data-ttu-id="47b77-313">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="47b77-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="47b77-314">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="47b77-314">Compose or Read</span></span>|
