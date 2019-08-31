---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: f06509e291325c635581d902d1f4f440bd255314
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696462"
---
# <a name="context"></a><span data-ttu-id="5756a-102">context</span><span class="sxs-lookup"><span data-stu-id="5756a-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="5756a-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="5756a-103">[Office](Office.md).context</span></span>

<span data-ttu-id="5756a-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="5756a-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5756a-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5756a-106">Requirements</span></span>

|<span data-ttu-id="5756a-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5756a-107">Requirement</span></span>| <span data-ttu-id="5756a-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="5756a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="5756a-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5756a-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5756a-110">1.0</span><span class="sxs-lookup"><span data-stu-id="5756a-110">1.0</span></span>|
|[<span data-ttu-id="5756a-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5756a-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5756a-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5756a-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="5756a-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="5756a-113">Members and methods</span></span>

| <span data-ttu-id="5756a-114">Membre</span><span class="sxs-lookup"><span data-stu-id="5756a-114">Member</span></span> | <span data-ttu-id="5756a-115">Type</span><span class="sxs-lookup"><span data-stu-id="5756a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="5756a-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="5756a-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="5756a-117">Member</span><span class="sxs-lookup"><span data-stu-id="5756a-117">Member</span></span> |
| [<span data-ttu-id="5756a-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="5756a-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="5756a-119">Member</span><span class="sxs-lookup"><span data-stu-id="5756a-119">Member</span></span> |
| [<span data-ttu-id="5756a-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="5756a-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="5756a-121">Membre</span><span class="sxs-lookup"><span data-stu-id="5756a-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="5756a-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="5756a-122">Namespaces</span></span>

<span data-ttu-id="5756a-123">[Mailbox](office.context.mailbox.md): permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="5756a-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="5756a-124">Members</span><span class="sxs-lookup"><span data-stu-id="5756a-124">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="5756a-125">displayLanguage: chaîne</span><span class="sxs-lookup"><span data-stu-id="5756a-125">displayLanguage: String</span></span>

<span data-ttu-id="5756a-126">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="5756a-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="5756a-127">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="5756a-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="5756a-128">Type</span><span class="sxs-lookup"><span data-stu-id="5756a-128">Type</span></span>

*   <span data-ttu-id="5756a-129">String</span><span class="sxs-lookup"><span data-stu-id="5756a-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5756a-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5756a-130">Requirements</span></span>

|<span data-ttu-id="5756a-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5756a-131">Requirement</span></span>| <span data-ttu-id="5756a-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="5756a-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="5756a-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5756a-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5756a-134">1.0</span><span class="sxs-lookup"><span data-stu-id="5756a-134">1.0</span></span>|
|[<span data-ttu-id="5756a-135">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5756a-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5756a-136">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5756a-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5756a-137">Exemple</span><span class="sxs-lookup"><span data-stu-id="5756a-137">Example</span></span>

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

#### <a name="officetheme-object"></a><span data-ttu-id="5756a-138">officeTheme: objet</span><span class="sxs-lookup"><span data-stu-id="5756a-138">officeTheme: Object</span></span>

<span data-ttu-id="5756a-139">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="5756a-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="5756a-140">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="5756a-140">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="5756a-141">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications hôtes Office.</span><span class="sxs-lookup"><span data-stu-id="5756a-141">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="5756a-142">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="5756a-142">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="5756a-143">Type</span><span class="sxs-lookup"><span data-stu-id="5756a-143">Type</span></span>

*   <span data-ttu-id="5756a-144">Objet</span><span class="sxs-lookup"><span data-stu-id="5756a-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="5756a-145">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="5756a-145">Properties:</span></span>

|<span data-ttu-id="5756a-146">Nom</span><span class="sxs-lookup"><span data-stu-id="5756a-146">Name</span></span>| <span data-ttu-id="5756a-147">Type</span><span class="sxs-lookup"><span data-stu-id="5756a-147">Type</span></span>| <span data-ttu-id="5756a-148">Description</span><span class="sxs-lookup"><span data-stu-id="5756a-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="5756a-149">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5756a-149">String</span></span>|<span data-ttu-id="5756a-150">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="5756a-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="5756a-151">Chaîne</span><span class="sxs-lookup"><span data-stu-id="5756a-151">String</span></span>|<span data-ttu-id="5756a-152">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="5756a-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="5756a-153">String</span><span class="sxs-lookup"><span data-stu-id="5756a-153">String</span></span>|<span data-ttu-id="5756a-154">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="5756a-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="5756a-155">String</span><span class="sxs-lookup"><span data-stu-id="5756a-155">String</span></span>|<span data-ttu-id="5756a-156">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="5756a-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="5756a-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5756a-157">Requirements</span></span>

|<span data-ttu-id="5756a-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5756a-158">Requirement</span></span>| <span data-ttu-id="5756a-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="5756a-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="5756a-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5756a-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5756a-161">Aperçu</span><span class="sxs-lookup"><span data-stu-id="5756a-161">Preview</span></span>|
|[<span data-ttu-id="5756a-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5756a-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5756a-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5756a-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5756a-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="5756a-164">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="5756a-165">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="5756a-165">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="5756a-166">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="5756a-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="5756a-167">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="5756a-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="5756a-168">Type</span><span class="sxs-lookup"><span data-stu-id="5756a-168">Type</span></span>

*   [<span data-ttu-id="5756a-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="5756a-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="5756a-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="5756a-170">Requirements</span></span>

|<span data-ttu-id="5756a-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="5756a-171">Requirement</span></span>| <span data-ttu-id="5756a-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="5756a-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="5756a-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="5756a-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="5756a-174">1.0</span><span class="sxs-lookup"><span data-stu-id="5756a-174">1.0</span></span>|
|[<span data-ttu-id="5756a-175">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="5756a-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="5756a-176">Restreinte</span><span class="sxs-lookup"><span data-stu-id="5756a-176">Restricted</span></span>|
|[<span data-ttu-id="5756a-177">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="5756a-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="5756a-178">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="5756a-178">Compose or Read</span></span>|
