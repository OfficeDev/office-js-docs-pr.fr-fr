---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629187"
---
# <a name="context"></a><span data-ttu-id="563c2-102">context</span><span class="sxs-lookup"><span data-stu-id="563c2-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="563c2-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="563c2-103">[Office](Office.md).context</span></span>

<span data-ttu-id="563c2-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="563c2-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="563c2-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-106">Requirements</span></span>

|<span data-ttu-id="563c2-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-107">Requirement</span></span>| <span data-ttu-id="563c2-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-110">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-110">1.0</span></span>|
|[<span data-ttu-id="563c2-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="563c2-113">Propriétés</span><span class="sxs-lookup"><span data-stu-id="563c2-113">Properties</span></span>

| <span data-ttu-id="563c2-114">Propriété</span><span class="sxs-lookup"><span data-stu-id="563c2-114">Property</span></span> | <span data-ttu-id="563c2-115">Modes</span><span class="sxs-lookup"><span data-stu-id="563c2-115">Modes</span></span> | <span data-ttu-id="563c2-116">Type de retour</span><span class="sxs-lookup"><span data-stu-id="563c2-116">Return type</span></span> | <span data-ttu-id="563c2-117">Minimale</span><span class="sxs-lookup"><span data-stu-id="563c2-117">Minimum</span></span><br><span data-ttu-id="563c2-118">ensemble de conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-118">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="563c2-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="563c2-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="563c2-120">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-120">Compose</span></span><br><span data-ttu-id="563c2-121">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-121">Read</span></span> | <span data-ttu-id="563c2-122">String</span><span class="sxs-lookup"><span data-stu-id="563c2-122">String</span></span> | <span data-ttu-id="563c2-123">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-123">1.0</span></span> |
| [<span data-ttu-id="563c2-124">Diagnostics</span><span class="sxs-lookup"><span data-stu-id="563c2-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="563c2-125">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-125">Compose</span></span><br><span data-ttu-id="563c2-126">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-126">Read</span></span> | [<span data-ttu-id="563c2-127">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="563c2-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation) | <span data-ttu-id="563c2-128">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-128">1.0</span></span> |
| [<span data-ttu-id="563c2-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="563c2-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="563c2-130">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-130">Compose</span></span><br><span data-ttu-id="563c2-131">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-131">Read</span></span> | <span data-ttu-id="563c2-132">String</span><span class="sxs-lookup"><span data-stu-id="563c2-132">String</span></span> | <span data-ttu-id="563c2-133">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-133">1.0</span></span> |
| [<span data-ttu-id="563c2-134">hote</span><span class="sxs-lookup"><span data-stu-id="563c2-134">host</span></span>](#host-hosttype) | <span data-ttu-id="563c2-135">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-135">Compose</span></span><br><span data-ttu-id="563c2-136">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-136">Read</span></span> | [<span data-ttu-id="563c2-137">HostType</span><span class="sxs-lookup"><span data-stu-id="563c2-137">HostType</span></span>](/javascript/api/office/office.hosttype) | <span data-ttu-id="563c2-138">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-138">1.0</span></span> |
| [<span data-ttu-id="563c2-139">officeTheme</span><span class="sxs-lookup"><span data-stu-id="563c2-139">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="563c2-140">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-140">Compose</span></span><br><span data-ttu-id="563c2-141">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-141">Read</span></span> | [<span data-ttu-id="563c2-142">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="563c2-142">OfficeTheme</span></span>](/javascript/api/office/office.officetheme) | <span data-ttu-id="563c2-143">Aperçu</span><span class="sxs-lookup"><span data-stu-id="563c2-143">Preview</span></span> |
| [<span data-ttu-id="563c2-144">plateforme</span><span class="sxs-lookup"><span data-stu-id="563c2-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="563c2-145">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-145">Compose</span></span><br><span data-ttu-id="563c2-146">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-146">Read</span></span> | [<span data-ttu-id="563c2-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="563c2-147">PlatformType</span></span>](/javascript/api/office/office.platformtype) | <span data-ttu-id="563c2-148">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-148">1.0</span></span> |
| [<span data-ttu-id="563c2-149">requise</span><span class="sxs-lookup"><span data-stu-id="563c2-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="563c2-150">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-150">Compose</span></span><br><span data-ttu-id="563c2-151">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-151">Read</span></span> | [<span data-ttu-id="563c2-152">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="563c2-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport) | <span data-ttu-id="563c2-153">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-153">1.0</span></span> |
| [<span data-ttu-id="563c2-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="563c2-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="563c2-155">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-155">Compose</span></span><br><span data-ttu-id="563c2-156">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-156">Read</span></span> | [<span data-ttu-id="563c2-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="563c2-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings) | <span data-ttu-id="563c2-158">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-158">1.0</span></span> |
| [<span data-ttu-id="563c2-159">ui</span><span class="sxs-lookup"><span data-stu-id="563c2-159">ui</span></span>](#ui-ui) | <span data-ttu-id="563c2-160">Composition</span><span class="sxs-lookup"><span data-stu-id="563c2-160">Compose</span></span><br><span data-ttu-id="563c2-161">Lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-161">Read</span></span> | [<span data-ttu-id="563c2-162">UI</span><span class="sxs-lookup"><span data-stu-id="563c2-162">UI</span></span>](/javascript/api/office/office.ui) | <span data-ttu-id="563c2-163">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-163">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="563c2-164">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="563c2-164">Namespaces</span></span>

<span data-ttu-id="563c2-165">[auth](/javascript/api/office/office.auth): fournit la prise en charge de l’authentification [unique (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span><span class="sxs-lookup"><span data-stu-id="563c2-165">[auth](/javascript/api/office/office.auth): Provides support for [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span></span>

<span data-ttu-id="563c2-166">[Mailbox](office.context.mailbox.md): permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="563c2-166">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

## <a name="property-details"></a><span data-ttu-id="563c2-167">Détails de la propriété</span><span class="sxs-lookup"><span data-stu-id="563c2-167">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="563c2-168">contentLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="563c2-168">contentLanguage: String</span></span>

<span data-ttu-id="563c2-169">Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.</span><span class="sxs-lookup"><span data-stu-id="563c2-169">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="563c2-170">La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="563c2-170">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-171">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-171">Type</span></span>

*   <span data-ttu-id="563c2-172">String</span><span class="sxs-lookup"><span data-stu-id="563c2-172">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="563c2-173">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-173">Requirements</span></span>

|<span data-ttu-id="563c2-174">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-174">Requirement</span></span>| <span data-ttu-id="563c2-175">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-176">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-177">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-177">1.0</span></span>|
|[<span data-ttu-id="563c2-178">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-179">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="563c2-180">Exemple</span><span class="sxs-lookup"><span data-stu-id="563c2-180">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="563c2-181">Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="563c2-181">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="563c2-182">Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="563c2-182">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-183">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-183">Type</span></span>

*   [<span data-ttu-id="563c2-184">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="563c2-184">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="563c2-185">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-185">Requirements</span></span>

|<span data-ttu-id="563c2-186">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-186">Requirement</span></span>| <span data-ttu-id="563c2-187">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-188">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-189">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-189">1.0</span></span>|
|[<span data-ttu-id="563c2-190">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-190">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-191">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-191">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="563c2-192">Exemple</span><span class="sxs-lookup"><span data-stu-id="563c2-192">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="563c2-193">displayLanguage : chaîne</span><span class="sxs-lookup"><span data-stu-id="563c2-193">displayLanguage: String</span></span>

<span data-ttu-id="563c2-194">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="563c2-194">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="563c2-195">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="563c2-195">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-196">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-196">Type</span></span>

*   <span data-ttu-id="563c2-197">String</span><span class="sxs-lookup"><span data-stu-id="563c2-197">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="563c2-198">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-198">Requirements</span></span>

|<span data-ttu-id="563c2-199">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-199">Requirement</span></span>| <span data-ttu-id="563c2-200">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-201">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-201">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-202">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-202">1.0</span></span>|
|[<span data-ttu-id="563c2-203">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-203">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-204">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="563c2-205">Exemple</span><span class="sxs-lookup"><span data-stu-id="563c2-205">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="563c2-206">hôte : [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="563c2-206">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="563c2-207">Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="563c2-207">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-208">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-208">Type</span></span>

*   [<span data-ttu-id="563c2-209">HostType</span><span class="sxs-lookup"><span data-stu-id="563c2-209">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="563c2-210">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-210">Requirements</span></span>

|<span data-ttu-id="563c2-211">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-211">Requirement</span></span>| <span data-ttu-id="563c2-212">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-213">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-214">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-214">1.0</span></span>|
|[<span data-ttu-id="563c2-215">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-216">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-216">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="563c2-217">Exemple</span><span class="sxs-lookup"><span data-stu-id="563c2-217">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a><span data-ttu-id="563c2-218">officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="563c2-218">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="563c2-219">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="563c2-219">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="563c2-220">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="563c2-220">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="563c2-221">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications hôtes Office.</span><span class="sxs-lookup"><span data-stu-id="563c2-221">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="563c2-222">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="563c2-222">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-223">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-223">Type</span></span>

*   [<span data-ttu-id="563c2-224">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="563c2-224">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="563c2-225">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="563c2-225">Properties:</span></span>

|<span data-ttu-id="563c2-226">Nom</span><span class="sxs-lookup"><span data-stu-id="563c2-226">Name</span></span>| <span data-ttu-id="563c2-227">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-227">Type</span></span>| <span data-ttu-id="563c2-228">Description</span><span class="sxs-lookup"><span data-stu-id="563c2-228">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="563c2-229">String</span><span class="sxs-lookup"><span data-stu-id="563c2-229">String</span></span>|<span data-ttu-id="563c2-230">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="563c2-230">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="563c2-231">String</span><span class="sxs-lookup"><span data-stu-id="563c2-231">String</span></span>|<span data-ttu-id="563c2-232">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="563c2-232">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="563c2-233">String</span><span class="sxs-lookup"><span data-stu-id="563c2-233">String</span></span>|<span data-ttu-id="563c2-234">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="563c2-234">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="563c2-235">String</span><span class="sxs-lookup"><span data-stu-id="563c2-235">String</span></span>|<span data-ttu-id="563c2-236">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="563c2-236">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="563c2-237">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-237">Requirements</span></span>

|<span data-ttu-id="563c2-238">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-238">Requirement</span></span>| <span data-ttu-id="563c2-239">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-240">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-241">Aperçu</span><span class="sxs-lookup"><span data-stu-id="563c2-241">Preview</span></span>|
|[<span data-ttu-id="563c2-242">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-243">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-243">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="563c2-244">Exemple</span><span class="sxs-lookup"><span data-stu-id="563c2-244">Example</span></span>

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

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="563c2-245">plateforme : [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="563c2-245">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="563c2-246">Fournit la plateforme sur laquelle le complément est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="563c2-246">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-247">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-247">Type</span></span>

*   [<span data-ttu-id="563c2-248">PlatformType</span><span class="sxs-lookup"><span data-stu-id="563c2-248">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="563c2-249">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-249">Requirements</span></span>

|<span data-ttu-id="563c2-250">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-250">Requirement</span></span>| <span data-ttu-id="563c2-251">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-252">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-252">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-253">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-253">1.0</span></span>|
|[<span data-ttu-id="563c2-254">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-255">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-255">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="563c2-256">Exemple</span><span class="sxs-lookup"><span data-stu-id="563c2-256">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="563c2-257">Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="563c2-257">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="563c2-258">Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.</span><span class="sxs-lookup"><span data-stu-id="563c2-258">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-259">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-259">Type</span></span>

*   [<span data-ttu-id="563c2-260">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="563c2-260">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="563c2-261">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-261">Requirements</span></span>

|<span data-ttu-id="563c2-262">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-262">Requirement</span></span>| <span data-ttu-id="563c2-263">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-264">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-265">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-265">1.0</span></span>|
|[<span data-ttu-id="563c2-266">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-266">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-267">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-267">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="563c2-268">Exemple</span><span class="sxs-lookup"><span data-stu-id="563c2-268">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="563c2-269">roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="563c2-269">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="563c2-270">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="563c2-270">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="563c2-271">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="563c2-271">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-272">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-272">Type</span></span>

*   [<span data-ttu-id="563c2-273">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="563c2-273">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="563c2-274">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-274">Requirements</span></span>

|<span data-ttu-id="563c2-275">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-275">Requirement</span></span>| <span data-ttu-id="563c2-276">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-277">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-278">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-278">1.0</span></span>|
|[<span data-ttu-id="563c2-279">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="563c2-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="563c2-280">Restreinte</span><span class="sxs-lookup"><span data-stu-id="563c2-280">Restricted</span></span>|
|[<span data-ttu-id="563c2-281">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-282">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-282">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="563c2-283">interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="563c2-283">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="563c2-284">Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.</span><span class="sxs-lookup"><span data-stu-id="563c2-284">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="563c2-285">Type</span><span class="sxs-lookup"><span data-stu-id="563c2-285">Type</span></span>

*   [<span data-ttu-id="563c2-286">UI</span><span class="sxs-lookup"><span data-stu-id="563c2-286">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="563c2-287">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="563c2-287">Requirements</span></span>

|<span data-ttu-id="563c2-288">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="563c2-288">Requirement</span></span>| <span data-ttu-id="563c2-289">Valeur</span><span class="sxs-lookup"><span data-stu-id="563c2-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="563c2-290">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="563c2-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="563c2-291">1.0</span><span class="sxs-lookup"><span data-stu-id="563c2-291">1.0</span></span>|
|[<span data-ttu-id="563c2-292">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="563c2-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="563c2-293">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="563c2-293">Compose or Read</span></span>|
