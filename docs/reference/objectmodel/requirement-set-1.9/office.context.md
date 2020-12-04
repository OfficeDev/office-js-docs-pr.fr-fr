---
title: Office. Context-ensemble de conditions requises 1,9
description: Membres de l’objet Office. Context disponibles pour les compléments Outlook utilisant l’ensemble de conditions requises de l’API de boîte aux lettres 1,9.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 3a8a9fe65ebf3c5a5ee63766f71dfce8e3f8d905
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570722"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="7a4c7-103">contexte (boîte aux lettres requise définie sur 1,9)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="7a4c7-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="7a4c7-104">[Office](office.md).context</span></span>

<span data-ttu-id="7a4c7-105">Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="7a4c7-106">Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="7a4c7-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a4c7-107">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-107">Requirements</span></span>

|<span data-ttu-id="7a4c7-108">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-108">Requirement</span></span>| <span data-ttu-id="7a4c7-109">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-110">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-111">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-111">1.1</span></span>|
|[<span data-ttu-id="7a4c7-112">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-113">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="7a4c7-114">Propriétés</span><span class="sxs-lookup"><span data-stu-id="7a4c7-114">Properties</span></span>

| <span data-ttu-id="7a4c7-115">Propriété</span><span class="sxs-lookup"><span data-stu-id="7a4c7-115">Property</span></span> | <span data-ttu-id="7a4c7-116">Modes</span><span class="sxs-lookup"><span data-stu-id="7a4c7-116">Modes</span></span> | <span data-ttu-id="7a4c7-117">Type de retour</span><span class="sxs-lookup"><span data-stu-id="7a4c7-117">Return type</span></span> | <span data-ttu-id="7a4c7-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="7a4c7-118">Minimum</span></span><br><span data-ttu-id="7a4c7-119">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7a4c7-120">auth</span><span class="sxs-lookup"><span data-stu-id="7a4c7-120">auth</span></span>](#auth-auth) | <span data-ttu-id="7a4c7-121">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-121">Compose</span></span><br><span data-ttu-id="7a4c7-122">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-122">Read</span></span> | [<span data-ttu-id="7a4c7-123">Auth</span><span class="sxs-lookup"><span data-stu-id="7a4c7-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-124">Ensembles 1,3</span><span class="sxs-lookup"><span data-stu-id="7a4c7-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="7a4c7-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="7a4c7-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="7a4c7-126">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-126">Compose</span></span><br><span data-ttu-id="7a4c7-127">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-127">Read</span></span> | <span data-ttu-id="7a4c7-128">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7a4c7-128">String</span></span> | [<span data-ttu-id="7a4c7-129">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7a4c7-130">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="7a4c7-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="7a4c7-131">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-131">Compose</span></span><br><span data-ttu-id="7a4c7-132">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-132">Read</span></span> | [<span data-ttu-id="7a4c7-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7a4c7-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7a4c7-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="7a4c7-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="7a4c7-136">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-136">Compose</span></span><br><span data-ttu-id="7a4c7-137">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-137">Read</span></span> | <span data-ttu-id="7a4c7-138">Chaîne</span><span class="sxs-lookup"><span data-stu-id="7a4c7-138">String</span></span> | [<span data-ttu-id="7a4c7-139">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7a4c7-140">hote</span><span class="sxs-lookup"><span data-stu-id="7a4c7-140">host</span></span>](#host-hosttype) | <span data-ttu-id="7a4c7-141">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-141">Compose</span></span><br><span data-ttu-id="7a4c7-142">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-142">Read</span></span> | [<span data-ttu-id="7a4c7-143">HostType</span><span class="sxs-lookup"><span data-stu-id="7a4c7-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-144">1,5</span><span class="sxs-lookup"><span data-stu-id="7a4c7-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7a4c7-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="7a4c7-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="7a4c7-146">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-146">Compose</span></span><br><span data-ttu-id="7a4c7-147">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-147">Read</span></span> | [<span data-ttu-id="7a4c7-148">Boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-149">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7a4c7-150">plateforme</span><span class="sxs-lookup"><span data-stu-id="7a4c7-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="7a4c7-151">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-151">Compose</span></span><br><span data-ttu-id="7a4c7-152">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-152">Read</span></span> | [<span data-ttu-id="7a4c7-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7a4c7-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-154">1,5</span><span class="sxs-lookup"><span data-stu-id="7a4c7-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="7a4c7-155">requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="7a4c7-156">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-156">Compose</span></span><br><span data-ttu-id="7a4c7-157">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-157">Read</span></span> | [<span data-ttu-id="7a4c7-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7a4c7-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-159">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7a4c7-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="7a4c7-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="7a4c7-161">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-161">Compose</span></span><br><span data-ttu-id="7a4c7-162">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-162">Read</span></span> | [<span data-ttu-id="7a4c7-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7a4c7-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-164">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7a4c7-165">ui</span><span class="sxs-lookup"><span data-stu-id="7a4c7-165">ui</span></span>](#ui-ui) | <span data-ttu-id="7a4c7-166">Composition</span><span class="sxs-lookup"><span data-stu-id="7a4c7-166">Compose</span></span><br><span data-ttu-id="7a4c7-167">Lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-167">Read</span></span> | [<span data-ttu-id="7a4c7-168">UI</span><span class="sxs-lookup"><span data-stu-id="7a4c7-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="7a4c7-169">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="7a4c7-170">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="7a4c7-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="7a4c7-171">AUTH : [auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="7a4c7-172">Prend en charge l’authentification [unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application Web du complément.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="7a4c7-173">Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="7a4c7-174">Voir l' [ensemble de conditions requises pour ensembles 1,3](../../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="7a4c7-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-175">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-175">Type</span></span>

*   [<span data-ttu-id="7a4c7-176">Auth</span><span class="sxs-lookup"><span data-stu-id="7a4c7-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="7a4c7-177">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-177">Requirements</span></span>

|<span data-ttu-id="7a4c7-178">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-178">Requirement</span></span>| <span data-ttu-id="7a4c7-179">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-180">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-181">S/O</span><span class="sxs-lookup"><span data-stu-id="7a4c7-181">N/A</span></span>|
|[<span data-ttu-id="7a4c7-182">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-183">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a4c7-184">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a4c7-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="7a4c7-185">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="7a4c7-185">contentLanguage: String</span></span>

<span data-ttu-id="7a4c7-186">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="7a4c7-187">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-188">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-188">Type</span></span>

*   <span data-ttu-id="7a4c7-189">String</span><span class="sxs-lookup"><span data-stu-id="7a4c7-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a4c7-190">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-190">Requirements</span></span>

|<span data-ttu-id="7a4c7-191">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-191">Requirement</span></span>| <span data-ttu-id="7a4c7-192">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-193">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-194">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-194">1.1</span></span>|
|[<span data-ttu-id="7a4c7-195">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-196">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a4c7-197">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a4c7-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="7a4c7-198">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="7a4c7-199">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-200">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-200">Type</span></span>

*   [<span data-ttu-id="7a4c7-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="7a4c7-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="7a4c7-202">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-202">Requirements</span></span>

|<span data-ttu-id="7a4c7-203">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-203">Requirement</span></span>| <span data-ttu-id="7a4c7-204">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-205">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-206">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-206">1.1</span></span>|
|[<span data-ttu-id="7a4c7-207">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-208">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a4c7-209">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a4c7-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="7a4c7-210">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="7a4c7-210">displayLanguage: String</span></span>

<span data-ttu-id="7a4c7-211">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="7a4c7-212">La `displayLanguage` valeur reflète le paramètre **langue d’affichage** actuel spécifié avec les **options de > de fichiers > langue** dans l’application cliente Office.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-213">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-213">Type</span></span>

*   <span data-ttu-id="7a4c7-214">String</span><span class="sxs-lookup"><span data-stu-id="7a4c7-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7a4c7-215">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-215">Requirements</span></span>

|<span data-ttu-id="7a4c7-216">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-216">Requirement</span></span>| <span data-ttu-id="7a4c7-217">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-218">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-219">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-219">1.1</span></span>|
|[<span data-ttu-id="7a4c7-220">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-221">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a4c7-222">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a4c7-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="7a4c7-223">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="7a4c7-224">Obtient l’application Office qui héberge le complément.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="7a4c7-225">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-226">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-226">Type</span></span>

*   [<span data-ttu-id="7a4c7-227">HostType</span><span class="sxs-lookup"><span data-stu-id="7a4c7-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="7a4c7-228">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-228">Requirements</span></span>

|<span data-ttu-id="7a4c7-229">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-229">Requirement</span></span>| <span data-ttu-id="7a4c7-230">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-231">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-232">1,5</span><span class="sxs-lookup"><span data-stu-id="7a4c7-232">1.5</span></span>|
|[<span data-ttu-id="7a4c7-233">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-234">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a4c7-235">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a4c7-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="7a4c7-236">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="7a4c7-237">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="7a4c7-238">Vous pouvez également utiliser la propriété [Office. Context. Diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-239">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-239">Type</span></span>

*   [<span data-ttu-id="7a4c7-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7a4c7-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="7a4c7-241">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-241">Requirements</span></span>

|<span data-ttu-id="7a4c7-242">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-242">Requirement</span></span>| <span data-ttu-id="7a4c7-243">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-244">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-245">1,5</span><span class="sxs-lookup"><span data-stu-id="7a4c7-245">1.5</span></span>|
|[<span data-ttu-id="7a4c7-246">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-247">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a4c7-248">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a4c7-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="7a4c7-249">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="7a4c7-250">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-251">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-251">Type</span></span>

*   [<span data-ttu-id="7a4c7-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="7a4c7-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="7a4c7-253">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-253">Requirements</span></span>

|<span data-ttu-id="7a4c7-254">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-254">Requirement</span></span>| <span data-ttu-id="7a4c7-255">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-256">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-257">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-257">1.1</span></span>|
|[<span data-ttu-id="7a4c7-258">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-259">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7a4c7-260">Exemple</span><span class="sxs-lookup"><span data-stu-id="7a4c7-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="7a4c7-261">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="7a4c7-262">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="7a4c7-263">L' `RoamingSettings` objet vous permet de stocker et d’accéder aux données d’un complément de messagerie qui est stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce complément lorsqu’il est exécuté à partir de n’importe quel client Outlook utilisé pour accéder à cette boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-264">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-264">Type</span></span>

*   [<span data-ttu-id="7a4c7-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7a4c7-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="7a4c7-266">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-266">Requirements</span></span>

|<span data-ttu-id="7a4c7-267">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-267">Requirement</span></span>| <span data-ttu-id="7a4c7-268">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-269">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-270">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-270">1.1</span></span>|
|[<span data-ttu-id="7a4c7-271">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="7a4c7-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="7a4c7-272">Restreinte</span><span class="sxs-lookup"><span data-stu-id="7a4c7-272">Restricted</span></span>|
|[<span data-ttu-id="7a4c7-273">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-274">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="7a4c7-275">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="7a4c7-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="7a4c7-276">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="7a4c7-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="7a4c7-277">Type</span><span class="sxs-lookup"><span data-stu-id="7a4c7-277">Type</span></span>

*   [<span data-ttu-id="7a4c7-278">UI</span><span class="sxs-lookup"><span data-stu-id="7a4c7-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="7a4c7-279">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="7a4c7-279">Requirements</span></span>

|<span data-ttu-id="7a4c7-280">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="7a4c7-280">Requirement</span></span>| <span data-ttu-id="7a4c7-281">Valeur</span><span class="sxs-lookup"><span data-stu-id="7a4c7-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="7a4c7-282">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="7a4c7-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7a4c7-283">1.1</span><span class="sxs-lookup"><span data-stu-id="7a4c7-283">1.1</span></span>|
|[<span data-ttu-id="7a4c7-284">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="7a4c7-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7a4c7-285">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="7a4c7-285">Compose or Read</span></span>|
