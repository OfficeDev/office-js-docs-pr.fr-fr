---
title: Office.context – prévisualisation d’ensemble de conditions requises
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 32b1948ace8dd90f92206708d3821fde2aae256b
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457704"
---
# <a name="context"></a><span data-ttu-id="6ae8b-102">context</span><span class="sxs-lookup"><span data-stu-id="6ae8b-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="6ae8b-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="6ae8b-103">[Office](Office.md).context</span></span>

<span data-ttu-id="6ae8b-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="6ae8b-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6ae8b-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6ae8b-106">Requirements</span></span>

|<span data-ttu-id="6ae8b-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6ae8b-107">Requirement</span></span>| <span data-ttu-id="6ae8b-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="6ae8b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ae8b-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6ae8b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ae8b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="6ae8b-110">1.0</span></span>|
|[<span data-ttu-id="6ae8b-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6ae8b-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ae8b-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6ae8b-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6ae8b-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="6ae8b-113">Members and methods</span></span>

| <span data-ttu-id="6ae8b-114">Membre</span><span class="sxs-lookup"><span data-stu-id="6ae8b-114">Member</span></span> | <span data-ttu-id="6ae8b-115">Type</span><span class="sxs-lookup"><span data-stu-id="6ae8b-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6ae8b-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="6ae8b-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="6ae8b-117">Membre</span><span class="sxs-lookup"><span data-stu-id="6ae8b-117">Member</span></span> |
| [<span data-ttu-id="6ae8b-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="6ae8b-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="6ae8b-119">Membre</span><span class="sxs-lookup"><span data-stu-id="6ae8b-119">Member</span></span> |
| [<span data-ttu-id="6ae8b-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="6ae8b-120">roamingSettings</span></span>](#roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings) | <span data-ttu-id="6ae8b-121">Membre</span><span class="sxs-lookup"><span data-stu-id="6ae8b-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="6ae8b-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="6ae8b-122">Namespaces</span></span>

<span data-ttu-id="6ae8b-123">[mailbox](office.context.mailbox.md)- Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="6ae8b-124">Membres</span><span class="sxs-lookup"><span data-stu-id="6ae8b-124">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="6ae8b-125">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="6ae8b-125">displayLanguage :String</span></span>

<span data-ttu-id="6ae8b-126">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="6ae8b-127">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="6ae8b-128">Type :</span><span class="sxs-lookup"><span data-stu-id="6ae8b-128">Type:</span></span>

*   <span data-ttu-id="6ae8b-129">Chaîne</span><span class="sxs-lookup"><span data-stu-id="6ae8b-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6ae8b-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6ae8b-130">Requirements</span></span>

|<span data-ttu-id="6ae8b-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6ae8b-131">Requirement</span></span>| <span data-ttu-id="6ae8b-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="6ae8b-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ae8b-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6ae8b-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ae8b-134">1.0</span><span class="sxs-lookup"><span data-stu-id="6ae8b-134">1.0</span></span>|
|[<span data-ttu-id="6ae8b-135">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6ae8b-135">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ae8b-136">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6ae8b-136">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6ae8b-137">Exemple</span><span class="sxs-lookup"><span data-stu-id="6ae8b-137">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="6ae8b-138">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="6ae8b-138">officeTheme :Object</span></span>

<span data-ttu-id="6ae8b-139">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="6ae8b-140">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-140">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="6ae8b-p102">À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier > Compte Office > Thème Office**, qui est appliqué à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6ae8b-143">Type :</span><span class="sxs-lookup"><span data-stu-id="6ae8b-143">Type:</span></span>

*   <span data-ttu-id="6ae8b-144">Objet</span><span class="sxs-lookup"><span data-stu-id="6ae8b-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="6ae8b-145">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="6ae8b-145">Properties:</span></span>

|<span data-ttu-id="6ae8b-146">Nom</span><span class="sxs-lookup"><span data-stu-id="6ae8b-146">Name</span></span>| <span data-ttu-id="6ae8b-147">Type</span><span class="sxs-lookup"><span data-stu-id="6ae8b-147">Type</span></span>| <span data-ttu-id="6ae8b-148">Description</span><span class="sxs-lookup"><span data-stu-id="6ae8b-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="6ae8b-149">String</span><span class="sxs-lookup"><span data-stu-id="6ae8b-149">String</span></span>|<span data-ttu-id="6ae8b-150">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="6ae8b-151">String</span><span class="sxs-lookup"><span data-stu-id="6ae8b-151">String</span></span>|<span data-ttu-id="6ae8b-152">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="6ae8b-153">String</span><span class="sxs-lookup"><span data-stu-id="6ae8b-153">String</span></span>|<span data-ttu-id="6ae8b-154">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="6ae8b-155">String</span><span class="sxs-lookup"><span data-stu-id="6ae8b-155">String</span></span>|<span data-ttu-id="6ae8b-156">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6ae8b-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6ae8b-157">Requirements</span></span>

|<span data-ttu-id="6ae8b-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6ae8b-158">Requirement</span></span>| <span data-ttu-id="6ae8b-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="6ae8b-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ae8b-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6ae8b-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ae8b-161">1.3</span><span class="sxs-lookup"><span data-stu-id="6ae8b-161">1.3</span></span>|
|[<span data-ttu-id="6ae8b-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6ae8b-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ae8b-163">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6ae8b-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="6ae8b-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="6ae8b-164">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="6ae8b-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="6ae8b-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="6ae8b-166">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6ae8b-167">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="6ae8b-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6ae8b-168">Type :</span><span class="sxs-lookup"><span data-stu-id="6ae8b-168">Type:</span></span>

*   [<span data-ttu-id="6ae8b-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6ae8b-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6ae8b-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="6ae8b-170">Requirements</span></span>

|<span data-ttu-id="6ae8b-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="6ae8b-171">Requirement</span></span>| <span data-ttu-id="6ae8b-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="6ae8b-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="6ae8b-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="6ae8b-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6ae8b-174">1.0</span><span class="sxs-lookup"><span data-stu-id="6ae8b-174">1.0</span></span>|
|[<span data-ttu-id="6ae8b-175">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="6ae8b-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="6ae8b-176">Restreinte</span><span class="sxs-lookup"><span data-stu-id="6ae8b-176">Restricted</span></span>|
|[<span data-ttu-id="6ae8b-177">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="6ae8b-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6ae8b-178">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="6ae8b-178">Compose or read</span></span>|