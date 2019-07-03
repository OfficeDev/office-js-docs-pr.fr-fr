---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 998e752cf2292eec4e05901325a0192e158c0b7f
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454831"
---
# <a name="context"></a><span data-ttu-id="778b0-102">context</span><span class="sxs-lookup"><span data-stu-id="778b0-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="778b0-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="778b0-103">[Office](Office.md).context</span></span>

<span data-ttu-id="778b0-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="778b0-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="778b0-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="778b0-106">Requirements</span></span>

|<span data-ttu-id="778b0-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="778b0-107">Requirement</span></span>| <span data-ttu-id="778b0-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="778b0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="778b0-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="778b0-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778b0-110">1.0</span><span class="sxs-lookup"><span data-stu-id="778b0-110">1.0</span></span>|
|[<span data-ttu-id="778b0-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="778b0-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778b0-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="778b0-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="778b0-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="778b0-113">Members and methods</span></span>

| <span data-ttu-id="778b0-114">Membre</span><span class="sxs-lookup"><span data-stu-id="778b0-114">Member</span></span> | <span data-ttu-id="778b0-115">Type</span><span class="sxs-lookup"><span data-stu-id="778b0-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="778b0-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="778b0-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="778b0-117">Member</span><span class="sxs-lookup"><span data-stu-id="778b0-117">Member</span></span> |
| [<span data-ttu-id="778b0-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="778b0-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="778b0-119">Member</span><span class="sxs-lookup"><span data-stu-id="778b0-119">Member</span></span> |
| [<span data-ttu-id="778b0-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="778b0-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="778b0-121">Membre</span><span class="sxs-lookup"><span data-stu-id="778b0-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="778b0-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="778b0-122">Namespaces</span></span>

<span data-ttu-id="778b0-123">[Mailbox](office.context.mailbox.md): permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="778b0-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="778b0-124">Members</span><span class="sxs-lookup"><span data-stu-id="778b0-124">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="778b0-125">displayLanguage: chaîne</span><span class="sxs-lookup"><span data-stu-id="778b0-125">displayLanguage: String</span></span>

<span data-ttu-id="778b0-126">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="778b0-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="778b0-127">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="778b0-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="778b0-128">Type</span><span class="sxs-lookup"><span data-stu-id="778b0-128">Type</span></span>

*   <span data-ttu-id="778b0-129">String</span><span class="sxs-lookup"><span data-stu-id="778b0-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="778b0-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="778b0-130">Requirements</span></span>

|<span data-ttu-id="778b0-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="778b0-131">Requirement</span></span>| <span data-ttu-id="778b0-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="778b0-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="778b0-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="778b0-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778b0-134">1.0</span><span class="sxs-lookup"><span data-stu-id="778b0-134">1.0</span></span>|
|[<span data-ttu-id="778b0-135">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="778b0-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778b0-136">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="778b0-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778b0-137">Exemple</span><span class="sxs-lookup"><span data-stu-id="778b0-137">Example</span></span>

```javascript
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

---
---

#### <a name="officetheme-object"></a><span data-ttu-id="778b0-138">officeTheme: objet</span><span class="sxs-lookup"><span data-stu-id="778b0-138">officeTheme: Object</span></span>

<span data-ttu-id="778b0-139">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="778b0-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="778b0-140">Ce membre est uniquement pris en charge dans Outlook sur Windows.</span><span class="sxs-lookup"><span data-stu-id="778b0-140">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="778b0-141">L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications hôtes Office.</span><span class="sxs-lookup"><span data-stu-id="778b0-141">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="778b0-142">Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="778b0-142">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="778b0-143">Type</span><span class="sxs-lookup"><span data-stu-id="778b0-143">Type</span></span>

*   <span data-ttu-id="778b0-144">Objet</span><span class="sxs-lookup"><span data-stu-id="778b0-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="778b0-145">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="778b0-145">Properties:</span></span>

|<span data-ttu-id="778b0-146">Nom</span><span class="sxs-lookup"><span data-stu-id="778b0-146">Name</span></span>| <span data-ttu-id="778b0-147">Type</span><span class="sxs-lookup"><span data-stu-id="778b0-147">Type</span></span>| <span data-ttu-id="778b0-148">Description</span><span class="sxs-lookup"><span data-stu-id="778b0-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="778b0-149">Chaîne</span><span class="sxs-lookup"><span data-stu-id="778b0-149">String</span></span>|<span data-ttu-id="778b0-150">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="778b0-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="778b0-151">Chaîne</span><span class="sxs-lookup"><span data-stu-id="778b0-151">String</span></span>|<span data-ttu-id="778b0-152">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="778b0-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="778b0-153">String</span><span class="sxs-lookup"><span data-stu-id="778b0-153">String</span></span>|<span data-ttu-id="778b0-154">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="778b0-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="778b0-155">String</span><span class="sxs-lookup"><span data-stu-id="778b0-155">String</span></span>|<span data-ttu-id="778b0-156">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="778b0-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="778b0-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="778b0-157">Requirements</span></span>

|<span data-ttu-id="778b0-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="778b0-158">Requirement</span></span>| <span data-ttu-id="778b0-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="778b0-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="778b0-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="778b0-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778b0-161">Aperçu</span><span class="sxs-lookup"><span data-stu-id="778b0-161">Preview</span></span>|
|[<span data-ttu-id="778b0-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="778b0-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778b0-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="778b0-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="778b0-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="778b0-164">Example</span></span>

```javascript
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

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="778b0-165">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="778b0-165">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="778b0-166">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="778b0-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="778b0-167">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="778b0-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="778b0-168">Type</span><span class="sxs-lookup"><span data-stu-id="778b0-168">Type</span></span>

*   [<span data-ttu-id="778b0-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="778b0-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="778b0-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="778b0-170">Requirements</span></span>

|<span data-ttu-id="778b0-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="778b0-171">Requirement</span></span>| <span data-ttu-id="778b0-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="778b0-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="778b0-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="778b0-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="778b0-174">1.0</span><span class="sxs-lookup"><span data-stu-id="778b0-174">1.0</span></span>|
|[<span data-ttu-id="778b0-175">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="778b0-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="778b0-176">Restreinte</span><span class="sxs-lookup"><span data-stu-id="778b0-176">Restricted</span></span>|
|[<span data-ttu-id="778b0-177">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="778b0-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="778b0-178">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="778b0-178">Compose or Read</span></span>|
