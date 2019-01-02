---
title: Office.context-ensemble de conditions requises 1.3
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 9652a1d4dcd48c437bb1156e9abc4c2aff575d59
ms.sourcegitcommit: 60fd8a3ac4a6d66cb9e075ce7e0cde3c888a5fe9
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/28/2018
ms.locfileid: "27457543"
---
# <a name="context"></a><span data-ttu-id="185f3-102">context</span><span class="sxs-lookup"><span data-stu-id="185f3-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="185f3-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="185f3-103">[Office](Office.md).context</span></span>

<span data-ttu-id="185f3-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="185f3-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="185f3-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185f3-106">Requirements</span></span>

|<span data-ttu-id="185f3-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185f3-107">Requirement</span></span>| <span data-ttu-id="185f3-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="185f3-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="185f3-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185f3-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185f3-110">1.0</span><span class="sxs-lookup"><span data-stu-id="185f3-110">1.0</span></span>|
|[<span data-ttu-id="185f3-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185f3-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="185f3-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="185f3-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="185f3-113">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="185f3-113">Namespaces</span></span>

<span data-ttu-id="185f3-114">[mailbox](office.context.mailbox.md)- Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="185f3-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="185f3-115">Membres</span><span class="sxs-lookup"><span data-stu-id="185f3-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="185f3-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="185f3-116">displayLanguage :String</span></span>

<span data-ttu-id="185f3-117">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="185f3-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="185f3-118">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="185f3-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="185f3-119">Type :</span><span class="sxs-lookup"><span data-stu-id="185f3-119">Type:</span></span>

*   <span data-ttu-id="185f3-120">Chaîne</span><span class="sxs-lookup"><span data-stu-id="185f3-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="185f3-121">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185f3-121">Requirements</span></span>

|<span data-ttu-id="185f3-122">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185f3-122">Requirement</span></span>| <span data-ttu-id="185f3-123">Valeur</span><span class="sxs-lookup"><span data-stu-id="185f3-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="185f3-124">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185f3-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185f3-125">1.0</span><span class="sxs-lookup"><span data-stu-id="185f3-125">1.0</span></span>|
|[<span data-ttu-id="185f3-126">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185f3-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="185f3-127">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="185f3-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="185f3-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="185f3-128">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="185f3-129">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="185f3-129">officeTheme :Object</span></span>

<span data-ttu-id="185f3-130">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="185f3-130">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="185f3-131">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="185f3-131">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="185f3-p102">À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier > Compte Office > Thème Office**, qui est appliqué à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="185f3-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="185f3-134">Type :</span><span class="sxs-lookup"><span data-stu-id="185f3-134">Type:</span></span>

*   <span data-ttu-id="185f3-135">Objet</span><span class="sxs-lookup"><span data-stu-id="185f3-135">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="185f3-136">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="185f3-136">Properties:</span></span>

|<span data-ttu-id="185f3-137">Nom</span><span class="sxs-lookup"><span data-stu-id="185f3-137">Name</span></span>| <span data-ttu-id="185f3-138">Type</span><span class="sxs-lookup"><span data-stu-id="185f3-138">Type</span></span>| <span data-ttu-id="185f3-139">Description</span><span class="sxs-lookup"><span data-stu-id="185f3-139">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="185f3-140">String</span><span class="sxs-lookup"><span data-stu-id="185f3-140">String</span></span>|<span data-ttu-id="185f3-141">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="185f3-141">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="185f3-142">String</span><span class="sxs-lookup"><span data-stu-id="185f3-142">String</span></span>|<span data-ttu-id="185f3-143">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="185f3-143">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="185f3-144">String</span><span class="sxs-lookup"><span data-stu-id="185f3-144">String</span></span>|<span data-ttu-id="185f3-145">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="185f3-145">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="185f3-146">String</span><span class="sxs-lookup"><span data-stu-id="185f3-146">String</span></span>|<span data-ttu-id="185f3-147">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="185f3-147">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="185f3-148">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185f3-148">Requirements</span></span>

|<span data-ttu-id="185f3-149">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185f3-149">Requirement</span></span>| <span data-ttu-id="185f3-150">Valeur</span><span class="sxs-lookup"><span data-stu-id="185f3-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="185f3-151">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185f3-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185f3-152">1.3</span><span class="sxs-lookup"><span data-stu-id="185f3-152">1.3</span></span>|
|[<span data-ttu-id="185f3-153">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185f3-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="185f3-154">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="185f3-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="185f3-155">Exemple</span><span class="sxs-lookup"><span data-stu-id="185f3-155">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook13officeroamingsettings"></a><span data-ttu-id="185f3-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="185f3-156">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_3/office.RoamingSettings)</span></span>

<span data-ttu-id="185f3-157">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="185f3-157">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="185f3-158">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="185f3-158">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="185f3-159">Type :</span><span class="sxs-lookup"><span data-stu-id="185f3-159">Type:</span></span>

*   [<span data-ttu-id="185f3-160">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="185f3-160">RoamingSettings</span></span>](/javascript/api/outlook_1_3/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="185f3-161">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="185f3-161">Requirements</span></span>

|<span data-ttu-id="185f3-162">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="185f3-162">Requirement</span></span>| <span data-ttu-id="185f3-163">Valeur</span><span class="sxs-lookup"><span data-stu-id="185f3-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="185f3-164">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="185f3-164">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="185f3-165">1.0</span><span class="sxs-lookup"><span data-stu-id="185f3-165">1.0</span></span>|
|[<span data-ttu-id="185f3-166">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="185f3-166">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="185f3-167">Restreinte</span><span class="sxs-lookup"><span data-stu-id="185f3-167">Restricted</span></span>|
|[<span data-ttu-id="185f3-168">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="185f3-168">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="185f3-169">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="185f3-169">Compose or read</span></span>|