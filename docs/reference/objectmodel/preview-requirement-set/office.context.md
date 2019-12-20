---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: b935d46b22e65fd293d6aae4b374cfeda9b34f5d
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814450"
---
# <a name="context"></a><span data-ttu-id="96b46-102">context</span><span class="sxs-lookup"><span data-stu-id="96b46-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="96b46-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="96b46-103">[Office](office.md).context</span></span>

<span data-ttu-id="96b46-104">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="96b46-105">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="96b46-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="96b46-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-106">Requirements</span></span>

|<span data-ttu-id="96b46-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-107">Requirement</span></span>| <span data-ttu-id="96b46-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-110">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-110">1.1</span></span>|
|[<span data-ttu-id="96b46-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="96b46-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="96b46-113">Properties</span></span>

| <span data-ttu-id="96b46-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="96b46-114">Property</span></span> | <span data-ttu-id="96b46-115">Modes</span><span class="sxs-lookup"><span data-stu-id="96b46-115">Modes</span></span> | <span data-ttu-id="96b46-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="96b46-116">Return type</span></span> | <span data-ttu-id="96b46-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="96b46-117">Minimum</span></span><br><span data-ttu-id="96b46-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="96b46-119">auth</span><span class="sxs-lookup"><span data-stu-id="96b46-119">auth</span></span>](#auth-auth) | <span data-ttu-id="96b46-120">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-120">Compose</span></span><br><span data-ttu-id="96b46-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-121">Read</span></span> | [<span data-ttu-id="96b46-122">Auth</span><span class="sxs-lookup"><span data-stu-id="96b46-122">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="96b46-123">Aperçu</span><span class="sxs-lookup"><span data-stu-id="96b46-123">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="96b46-124">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="96b46-124">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="96b46-125">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-125">Compose</span></span><br><span data-ttu-id="96b46-126">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-126">Read</span></span> | <span data-ttu-id="96b46-127">String</span><span class="sxs-lookup"><span data-stu-id="96b46-127">String</span></span> | [<span data-ttu-id="96b46-128">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-129">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="96b46-129">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="96b46-130">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-130">Compose</span></span><br><span data-ttu-id="96b46-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-131">Read</span></span> | [<span data-ttu-id="96b46-132">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="96b46-132">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="96b46-133">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-134">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="96b46-134">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="96b46-135">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-135">Compose</span></span><br><span data-ttu-id="96b46-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-136">Read</span></span> | <span data-ttu-id="96b46-137">String</span><span class="sxs-lookup"><span data-stu-id="96b46-137">String</span></span> | [<span data-ttu-id="96b46-138">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-139">hote</span><span class="sxs-lookup"><span data-stu-id="96b46-139">host</span></span>](#host-hosttype) | <span data-ttu-id="96b46-140">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-140">Compose</span></span><br><span data-ttu-id="96b46-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-141">Read</span></span> | [<span data-ttu-id="96b46-142">HostType</span><span class="sxs-lookup"><span data-stu-id="96b46-142">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="96b46-143">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-144">mailbox</span><span class="sxs-lookup"><span data-stu-id="96b46-144">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="96b46-145">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-145">Compose</span></span><br><span data-ttu-id="96b46-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-146">Read</span></span> | [<span data-ttu-id="96b46-147">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-147">Mailbox</span></span>](/javascript/api/office/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="96b46-148">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-149">officeTheme</span><span class="sxs-lookup"><span data-stu-id="96b46-149">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="96b46-150">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-150">Compose</span></span><br><span data-ttu-id="96b46-151">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-151">Read</span></span> | [<span data-ttu-id="96b46-152">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="96b46-152">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="96b46-153">Aperçu</span><span class="sxs-lookup"><span data-stu-id="96b46-153">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="96b46-154">plateforme</span><span class="sxs-lookup"><span data-stu-id="96b46-154">platform</span></span>](#platform-platformtype) | <span data-ttu-id="96b46-155">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-155">Compose</span></span><br><span data-ttu-id="96b46-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-156">Read</span></span> | [<span data-ttu-id="96b46-157">PlatformType</span><span class="sxs-lookup"><span data-stu-id="96b46-157">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="96b46-158">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-159">requise</span><span class="sxs-lookup"><span data-stu-id="96b46-159">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="96b46-160">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-160">Compose</span></span><br><span data-ttu-id="96b46-161">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-161">Read</span></span> | [<span data-ttu-id="96b46-162">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="96b46-162">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="96b46-163">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-164">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="96b46-164">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="96b46-165">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-165">Compose</span></span><br><span data-ttu-id="96b46-166">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-166">Read</span></span> | [<span data-ttu-id="96b46-167">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="96b46-167">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="96b46-168">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-168">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96b46-169">ui</span><span class="sxs-lookup"><span data-stu-id="96b46-169">ui</span></span>](#ui-ui) | <span data-ttu-id="96b46-170">Composition</span><span class="sxs-lookup"><span data-stu-id="96b46-170">Compose</span></span><br><span data-ttu-id="96b46-171">Lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-171">Read</span></span> | [<span data-ttu-id="96b46-172">UI</span><span class="sxs-lookup"><span data-stu-id="96b46-172">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="96b46-173">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-173">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="96b46-174">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="96b46-174">Property details</span></span>

#### <a name="auth-authjavascriptapiofficeofficeauth"></a><span data-ttu-id="96b46-175">AUTH : [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="96b46-175">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="96b46-176">Prend en charge l’authentification [unique (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token) en fournissant une méthode qui permet à l’hôte Office d’obtenir un jeton d’accès à l’application Web du complément.</span><span class="sxs-lookup"><span data-stu-id="96b46-176">Supports [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="96b46-177">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="96b46-177">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-178">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-178">Type</span></span>

*   [<span data-ttu-id="96b46-179">Auth</span><span class="sxs-lookup"><span data-stu-id="96b46-179">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="96b46-180">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-180">Requirements</span></span>

|<span data-ttu-id="96b46-181">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-181">Requirement</span></span>| <span data-ttu-id="96b46-182">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-183">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-183">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-184">Aperçu</span><span class="sxs-lookup"><span data-stu-id="96b46-184">Preview</span></span>|
|[<span data-ttu-id="96b46-185">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-186">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-186">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-187">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-187">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="96b46-188">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="96b46-188">contentLanguage: String</span></span>

<span data-ttu-id="96b46-189">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="96b46-189">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="96b46-190">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-190">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-191">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-191">Type</span></span>

*   <span data-ttu-id="96b46-192">String</span><span class="sxs-lookup"><span data-stu-id="96b46-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96b46-193">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-193">Requirements</span></span>

|<span data-ttu-id="96b46-194">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-194">Requirement</span></span>| <span data-ttu-id="96b46-195">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-196">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-196">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-197">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-197">1.1</span></span>|
|[<span data-ttu-id="96b46-198">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-199">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-199">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-200">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-200">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="96b46-201">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="96b46-201">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="96b46-202">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="96b46-202">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-203">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-203">Type</span></span>

*   [<span data-ttu-id="96b46-204">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="96b46-204">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="96b46-205">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-205">Requirements</span></span>

|<span data-ttu-id="96b46-206">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-206">Requirement</span></span>| <span data-ttu-id="96b46-207">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-208">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-208">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-209">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-209">1.1</span></span>|
|[<span data-ttu-id="96b46-210">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-211">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-211">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-212">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-212">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="96b46-213">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="96b46-213">displayLanguage: String</span></span>

<span data-ttu-id="96b46-214">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-214">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="96b46-215">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-216">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-216">Type</span></span>

*   <span data-ttu-id="96b46-217">String</span><span class="sxs-lookup"><span data-stu-id="96b46-217">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96b46-218">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-218">Requirements</span></span>

|<span data-ttu-id="96b46-219">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-219">Requirement</span></span>| <span data-ttu-id="96b46-220">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-221">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-222">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-222">1.1</span></span>|
|[<span data-ttu-id="96b46-223">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-223">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-224">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-225">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-225">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="96b46-226">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="96b46-226">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="96b46-227">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="96b46-227">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-228">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-228">Type</span></span>

*   [<span data-ttu-id="96b46-229">HostType</span><span class="sxs-lookup"><span data-stu-id="96b46-229">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="96b46-230">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-230">Requirements</span></span>

|<span data-ttu-id="96b46-231">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-231">Requirement</span></span>| <span data-ttu-id="96b46-232">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-233">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-234">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-234">1.1</span></span>|
|[<span data-ttu-id="96b46-235">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-235">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-236">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-236">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-237">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-237">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a><span data-ttu-id="96b46-238">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="96b46-238">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="96b46-239">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-239">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="96b46-240">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="96b46-240">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="96b46-241">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications hôtes Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-241">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="96b46-242">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-243">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-243">Type</span></span>

*   [<span data-ttu-id="96b46-244">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="96b46-244">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="96b46-245">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="96b46-245">Properties:</span></span>

|<span data-ttu-id="96b46-246">Nom</span><span class="sxs-lookup"><span data-stu-id="96b46-246">Name</span></span>| <span data-ttu-id="96b46-247">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-247">Type</span></span>| <span data-ttu-id="96b46-248">Description</span><span class="sxs-lookup"><span data-stu-id="96b46-248">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="96b46-249">String</span><span class="sxs-lookup"><span data-stu-id="96b46-249">String</span></span>|<span data-ttu-id="96b46-250">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="96b46-250">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="96b46-251">String</span><span class="sxs-lookup"><span data-stu-id="96b46-251">String</span></span>|<span data-ttu-id="96b46-252">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="96b46-252">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="96b46-253">String</span><span class="sxs-lookup"><span data-stu-id="96b46-253">String</span></span>|<span data-ttu-id="96b46-254">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="96b46-254">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="96b46-255">String</span><span class="sxs-lookup"><span data-stu-id="96b46-255">String</span></span>|<span data-ttu-id="96b46-256">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="96b46-256">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="96b46-257">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-257">Requirements</span></span>

|<span data-ttu-id="96b46-258">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-258">Requirement</span></span>| <span data-ttu-id="96b46-259">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-260">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-261">Aperçu</span><span class="sxs-lookup"><span data-stu-id="96b46-261">Preview</span></span>|
|[<span data-ttu-id="96b46-262">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-262">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-263">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-263">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-264">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-264">Example</span></span>

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

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="96b46-265">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="96b46-265">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="96b46-266">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="96b46-266">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-267">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-267">Type</span></span>

*   [<span data-ttu-id="96b46-268">PlatformType</span><span class="sxs-lookup"><span data-stu-id="96b46-268">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="96b46-269">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-269">Requirements</span></span>

|<span data-ttu-id="96b46-270">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-270">Requirement</span></span>| <span data-ttu-id="96b46-271">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-272">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-272">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-273">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-273">1.1</span></span>|
|[<span data-ttu-id="96b46-274">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-274">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-275">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-275">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-276">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-276">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="96b46-277">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="96b46-277">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="96b46-278">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="96b46-278">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-279">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-279">Type</span></span>

*   [<span data-ttu-id="96b46-280">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="96b46-280">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="96b46-281">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-281">Requirements</span></span>

|<span data-ttu-id="96b46-282">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-282">Requirement</span></span>| <span data-ttu-id="96b46-283">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-284">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-284">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-285">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-285">1.1</span></span>|
|[<span data-ttu-id="96b46-286">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-286">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-287">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96b46-288">Exemple</span><span class="sxs-lookup"><span data-stu-id="96b46-288">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="96b46-289">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="96b46-289">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="96b46-290">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="96b46-290">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="96b46-291">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="96b46-291">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-292">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-292">Type</span></span>

*   [<span data-ttu-id="96b46-293">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="96b46-293">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="96b46-294">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-294">Requirements</span></span>

|<span data-ttu-id="96b46-295">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-295">Requirement</span></span>| <span data-ttu-id="96b46-296">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-297">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-297">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-298">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-298">1.1</span></span>|
|[<span data-ttu-id="96b46-299">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="96b46-299">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="96b46-300">Restreinte</span><span class="sxs-lookup"><span data-stu-id="96b46-300">Restricted</span></span>|
|[<span data-ttu-id="96b46-301">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-301">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-302">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-302">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="96b46-303">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="96b46-303">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="96b46-304">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="96b46-304">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="96b46-305">Type</span><span class="sxs-lookup"><span data-stu-id="96b46-305">Type</span></span>

*   [<span data-ttu-id="96b46-306">UI</span><span class="sxs-lookup"><span data-stu-id="96b46-306">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="96b46-307">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96b46-307">Requirements</span></span>

|<span data-ttu-id="96b46-308">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96b46-308">Requirement</span></span>| <span data-ttu-id="96b46-309">Valeur</span><span class="sxs-lookup"><span data-stu-id="96b46-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="96b46-310">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96b46-310">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96b46-311">1.1</span><span class="sxs-lookup"><span data-stu-id="96b46-311">1.1</span></span>|
|[<span data-ttu-id="96b46-312">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96b46-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="96b46-313">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96b46-313">Compose or Read</span></span>|
