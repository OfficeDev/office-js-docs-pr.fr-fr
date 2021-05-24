---
title: Office.context - ensemble de conditions requises 1.9
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.9.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: f45eec7ce638f4bbb97ad4be9f2ba089905c631d
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590518"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="96bc3-103">contexte (ensemble de conditions requises de boîte aux lettres 1.9)</span><span class="sxs-lookup"><span data-stu-id="96bc3-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="96bc3-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="96bc3-104">[Office](office.md).context</span></span>

<span data-ttu-id="96bc3-105">Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications.</span><span class="sxs-lookup"><span data-stu-id="96bc3-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="96bc3-106">Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir la liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="96bc3-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="96bc3-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-107">Requirements</span></span>

|<span data-ttu-id="96bc3-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-108">Requirement</span></span>| <span data-ttu-id="96bc3-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-111">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-111">1.1</span></span>|
|[<span data-ttu-id="96bc3-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="96bc3-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="96bc3-114">Properties</span></span>

| <span data-ttu-id="96bc3-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="96bc3-115">Property</span></span> | <span data-ttu-id="96bc3-116">Modes</span><span class="sxs-lookup"><span data-stu-id="96bc3-116">Modes</span></span> | <span data-ttu-id="96bc3-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="96bc3-117">Return type</span></span> | <span data-ttu-id="96bc3-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="96bc3-118">Minimum</span></span><br><span data-ttu-id="96bc3-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="96bc3-120">auth</span><span class="sxs-lookup"><span data-stu-id="96bc3-120">auth</span></span>](#auth-auth) | <span data-ttu-id="96bc3-121">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-121">Compose</span></span><br><span data-ttu-id="96bc3-122">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-122">Read</span></span> | [<span data-ttu-id="96bc3-123">Auth</span><span class="sxs-lookup"><span data-stu-id="96bc3-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="96bc3-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="96bc3-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="96bc3-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="96bc3-126">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-126">Compose</span></span><br><span data-ttu-id="96bc3-127">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-127">Read</span></span> | <span data-ttu-id="96bc3-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="96bc3-128">String</span></span> | [<span data-ttu-id="96bc3-129">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96bc3-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="96bc3-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="96bc3-131">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-131">Compose</span></span><br><span data-ttu-id="96bc3-132">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-132">Read</span></span> | [<span data-ttu-id="96bc3-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="96bc3-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-134">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96bc3-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="96bc3-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="96bc3-136">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-136">Compose</span></span><br><span data-ttu-id="96bc3-137">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-137">Read</span></span> | <span data-ttu-id="96bc3-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="96bc3-138">String</span></span> | [<span data-ttu-id="96bc3-139">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96bc3-140">host</span><span class="sxs-lookup"><span data-stu-id="96bc3-140">host</span></span>](#host-hosttype) | <span data-ttu-id="96bc3-141">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-141">Compose</span></span><br><span data-ttu-id="96bc3-142">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-142">Read</span></span> | [<span data-ttu-id="96bc3-143">HostType</span><span class="sxs-lookup"><span data-stu-id="96bc3-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-144">1.5</span><span class="sxs-lookup"><span data-stu-id="96bc3-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="96bc3-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="96bc3-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="96bc3-146">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-146">Compose</span></span><br><span data-ttu-id="96bc3-147">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-147">Read</span></span> | [<span data-ttu-id="96bc3-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-149">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96bc3-150">platform</span><span class="sxs-lookup"><span data-stu-id="96bc3-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="96bc3-151">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-151">Compose</span></span><br><span data-ttu-id="96bc3-152">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-152">Read</span></span> | [<span data-ttu-id="96bc3-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="96bc3-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-154">1.5</span><span class="sxs-lookup"><span data-stu-id="96bc3-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="96bc3-155">requirements</span><span class="sxs-lookup"><span data-stu-id="96bc3-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="96bc3-156">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-156">Compose</span></span><br><span data-ttu-id="96bc3-157">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-157">Read</span></span> | [<span data-ttu-id="96bc3-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="96bc3-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-159">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96bc3-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="96bc3-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="96bc3-161">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-161">Compose</span></span><br><span data-ttu-id="96bc3-162">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-162">Read</span></span> | [<span data-ttu-id="96bc3-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="96bc3-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-164">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="96bc3-165">ui</span><span class="sxs-lookup"><span data-stu-id="96bc3-165">ui</span></span>](#ui-ui) | <span data-ttu-id="96bc3-166">Composition</span><span class="sxs-lookup"><span data-stu-id="96bc3-166">Compose</span></span><br><span data-ttu-id="96bc3-167">Lire</span><span class="sxs-lookup"><span data-stu-id="96bc3-167">Read</span></span> | [<span data-ttu-id="96bc3-168">UI</span><span class="sxs-lookup"><span data-stu-id="96bc3-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="96bc3-169">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="96bc3-170">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="96bc3-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="96bc3-171">auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="96bc3-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="96bc3-172">Prend en charge l' [sign-on unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application web du module.</span><span class="sxs-lookup"><span data-stu-id="96bc3-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="96bc3-173">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="96bc3-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="96bc3-174">Voir [l’ensemble de conditions requises IdentityAPI 1.3.](../../requirement-sets/identity-api-requirement-sets.md)</span><span class="sxs-lookup"><span data-stu-id="96bc3-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-175">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-175">Type</span></span>

*   [<span data-ttu-id="96bc3-176">Auth</span><span class="sxs-lookup"><span data-stu-id="96bc3-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="96bc3-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-177">Requirements</span></span>

|<span data-ttu-id="96bc3-178">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-178">Requirement</span></span>| <span data-ttu-id="96bc3-179">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-180">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-181">S/O</span><span class="sxs-lookup"><span data-stu-id="96bc3-181">N/A</span></span>|
|[<span data-ttu-id="96bc3-182">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-183">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96bc3-184">Exemple</span><span class="sxs-lookup"><span data-stu-id="96bc3-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="96bc3-185">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="96bc3-185">contentLanguage: String</span></span>

<span data-ttu-id="96bc3-186">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="96bc3-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="96bc3-187">La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options de > langue dans l Office `contentLanguage` application cliente.  </span><span class="sxs-lookup"><span data-stu-id="96bc3-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-188">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-188">Type</span></span>

*   <span data-ttu-id="96bc3-189">String</span><span class="sxs-lookup"><span data-stu-id="96bc3-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96bc3-190">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-190">Requirements</span></span>

|<span data-ttu-id="96bc3-191">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-191">Requirement</span></span>| <span data-ttu-id="96bc3-192">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-193">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-194">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-194">1.1</span></span>|
|[<span data-ttu-id="96bc3-195">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-196">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96bc3-197">Exemple</span><span class="sxs-lookup"><span data-stu-id="96bc3-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="96bc3-198">diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="96bc3-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="96bc3-199">Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="96bc3-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-200">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-200">Type</span></span>

*   [<span data-ttu-id="96bc3-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="96bc3-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="96bc3-202">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-202">Requirements</span></span>

|<span data-ttu-id="96bc3-203">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-203">Requirement</span></span>| <span data-ttu-id="96bc3-204">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-205">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-206">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-206">1.1</span></span>|
|[<span data-ttu-id="96bc3-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96bc3-209">Exemple</span><span class="sxs-lookup"><span data-stu-id="96bc3-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="96bc3-210">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="96bc3-210">displayLanguage: String</span></span>

<span data-ttu-id="96bc3-211">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.</span><span class="sxs-lookup"><span data-stu-id="96bc3-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="96bc3-212">La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office..  </span><span class="sxs-lookup"><span data-stu-id="96bc3-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-213">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-213">Type</span></span>

*   <span data-ttu-id="96bc3-214">String</span><span class="sxs-lookup"><span data-stu-id="96bc3-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="96bc3-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-215">Requirements</span></span>

|<span data-ttu-id="96bc3-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-216">Requirement</span></span>| <span data-ttu-id="96bc3-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-219">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-219">1.1</span></span>|
|[<span data-ttu-id="96bc3-220">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-221">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96bc3-222">Exemple</span><span class="sxs-lookup"><span data-stu-id="96bc3-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="96bc3-223">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="96bc3-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="96bc3-224">Obtient Office application qui héberge le module.</span><span class="sxs-lookup"><span data-stu-id="96bc3-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="96bc3-225">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="96bc3-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-226">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-226">Type</span></span>

*   [<span data-ttu-id="96bc3-227">HostType</span><span class="sxs-lookup"><span data-stu-id="96bc3-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="96bc3-228">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-228">Requirements</span></span>

|<span data-ttu-id="96bc3-229">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-229">Requirement</span></span>| <span data-ttu-id="96bc3-230">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-231">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-232">1,5</span><span class="sxs-lookup"><span data-stu-id="96bc3-232">1.5</span></span>|
|[<span data-ttu-id="96bc3-233">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-234">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96bc3-235">Exemple</span><span class="sxs-lookup"><span data-stu-id="96bc3-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="96bc3-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="96bc3-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="96bc3-237">Fournit la plateforme sur laquelle le module est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="96bc3-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="96bc3-238">Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="96bc3-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-239">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-239">Type</span></span>

*   [<span data-ttu-id="96bc3-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="96bc3-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="96bc3-241">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-241">Requirements</span></span>

|<span data-ttu-id="96bc3-242">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-242">Requirement</span></span>| <span data-ttu-id="96bc3-243">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-244">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-245">1,5</span><span class="sxs-lookup"><span data-stu-id="96bc3-245">1.5</span></span>|
|[<span data-ttu-id="96bc3-246">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-247">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96bc3-248">Exemple</span><span class="sxs-lookup"><span data-stu-id="96bc3-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="96bc3-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="96bc3-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="96bc3-250">Fournit une méthode pour déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="96bc3-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-251">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-251">Type</span></span>

*   [<span data-ttu-id="96bc3-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="96bc3-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="96bc3-253">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-253">Requirements</span></span>

|<span data-ttu-id="96bc3-254">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-254">Requirement</span></span>| <span data-ttu-id="96bc3-255">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-256">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-257">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-257">1.1</span></span>|
|[<span data-ttu-id="96bc3-258">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-259">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="96bc3-260">Exemple</span><span class="sxs-lookup"><span data-stu-id="96bc3-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="96bc3-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="96bc3-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="96bc3-262">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="96bc3-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="96bc3-263">L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.</span><span class="sxs-lookup"><span data-stu-id="96bc3-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-264">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-264">Type</span></span>

*   [<span data-ttu-id="96bc3-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="96bc3-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="96bc3-266">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-266">Requirements</span></span>

|<span data-ttu-id="96bc3-267">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-267">Requirement</span></span>| <span data-ttu-id="96bc3-268">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-269">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-270">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-270">1.1</span></span>|
|[<span data-ttu-id="96bc3-271">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="96bc3-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="96bc3-272">Restreinte</span><span class="sxs-lookup"><span data-stu-id="96bc3-272">Restricted</span></span>|
|[<span data-ttu-id="96bc3-273">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-274">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="96bc3-275">Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="96bc3-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="96bc3-276">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.</span><span class="sxs-lookup"><span data-stu-id="96bc3-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="96bc3-277">Type</span><span class="sxs-lookup"><span data-stu-id="96bc3-277">Type</span></span>

*   [<span data-ttu-id="96bc3-278">UI</span><span class="sxs-lookup"><span data-stu-id="96bc3-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="96bc3-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="96bc3-279">Requirements</span></span>

|<span data-ttu-id="96bc3-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="96bc3-280">Requirement</span></span>| <span data-ttu-id="96bc3-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="96bc3-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="96bc3-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="96bc3-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="96bc3-283">1.1</span><span class="sxs-lookup"><span data-stu-id="96bc3-283">1.1</span></span>|
|[<span data-ttu-id="96bc3-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="96bc3-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="96bc3-285">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="96bc3-285">Compose or Read</span></span>|
