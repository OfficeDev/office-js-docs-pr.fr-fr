---
title: Office. Context-ensemble de conditions requises 1,7
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: dddf0035f52daadc926ca5a707383730a97c1002
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838472"
---
# <a name="context"></a><span data-ttu-id="4ed9f-102">context</span><span class="sxs-lookup"><span data-stu-id="4ed9f-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="4ed9f-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="4ed9f-103">[Office](Office.md).context</span></span>

<span data-ttu-id="4ed9f-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="4ed9f-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ed9f-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ed9f-106">Requirements</span></span>

|<span data-ttu-id="4ed9f-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ed9f-107">Requirement</span></span>| <span data-ttu-id="4ed9f-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ed9f-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ed9f-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ed9f-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ed9f-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4ed9f-110">1.0</span></span>|
|[<span data-ttu-id="4ed9f-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ed9f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4ed9f-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ed9f-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="4ed9f-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="4ed9f-113">Members and methods</span></span>

| <span data-ttu-id="4ed9f-114">Membre</span><span class="sxs-lookup"><span data-stu-id="4ed9f-114">Member</span></span> | <span data-ttu-id="4ed9f-115">Type</span><span class="sxs-lookup"><span data-stu-id="4ed9f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="4ed9f-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="4ed9f-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="4ed9f-117">Member</span><span class="sxs-lookup"><span data-stu-id="4ed9f-117">Member</span></span> |
| [<span data-ttu-id="4ed9f-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="4ed9f-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="4ed9f-119">Member</span><span class="sxs-lookup"><span data-stu-id="4ed9f-119">Member</span></span> |
| [<span data-ttu-id="4ed9f-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="4ed9f-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="4ed9f-121">Membre</span><span class="sxs-lookup"><span data-stu-id="4ed9f-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="4ed9f-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="4ed9f-122">Namespaces</span></span>

<span data-ttu-id="4ed9f-123">[Mailbox](office.context.mailbox.md): permet d'accéder au modèle d'objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="4ed9f-124">Membres</span><span class="sxs-lookup"><span data-stu-id="4ed9f-124">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="4ed9f-125">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="4ed9f-125">displayLanguage :String</span></span>

<span data-ttu-id="4ed9f-126">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="4ed9f-127">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="4ed9f-128">Type</span><span class="sxs-lookup"><span data-stu-id="4ed9f-128">Type</span></span>

*   <span data-ttu-id="4ed9f-129">String</span><span class="sxs-lookup"><span data-stu-id="4ed9f-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4ed9f-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ed9f-130">Requirements</span></span>

|<span data-ttu-id="4ed9f-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ed9f-131">Requirement</span></span>| <span data-ttu-id="4ed9f-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ed9f-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ed9f-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ed9f-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ed9f-134">1.0</span><span class="sxs-lookup"><span data-stu-id="4ed9f-134">1.0</span></span>|
|[<span data-ttu-id="4ed9f-135">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ed9f-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4ed9f-136">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ed9f-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ed9f-137">Exemple</span><span class="sxs-lookup"><span data-stu-id="4ed9f-137">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="4ed9f-138">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="4ed9f-138">officeTheme :Object</span></span>

<span data-ttu-id="4ed9f-139">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="4ed9f-140">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-140">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4ed9f-p102">À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier > Compte Office > Thème Office**, qui est appliqué à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="4ed9f-143">Type</span><span class="sxs-lookup"><span data-stu-id="4ed9f-143">Type</span></span>

*   <span data-ttu-id="4ed9f-144">Objet</span><span class="sxs-lookup"><span data-stu-id="4ed9f-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="4ed9f-145">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="4ed9f-145">Properties:</span></span>

|<span data-ttu-id="4ed9f-146">Nom</span><span class="sxs-lookup"><span data-stu-id="4ed9f-146">Name</span></span>| <span data-ttu-id="4ed9f-147">Type</span><span class="sxs-lookup"><span data-stu-id="4ed9f-147">Type</span></span>| <span data-ttu-id="4ed9f-148">Description</span><span class="sxs-lookup"><span data-stu-id="4ed9f-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="4ed9f-149">String</span><span class="sxs-lookup"><span data-stu-id="4ed9f-149">String</span></span>|<span data-ttu-id="4ed9f-150">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="4ed9f-151">String</span><span class="sxs-lookup"><span data-stu-id="4ed9f-151">String</span></span>|<span data-ttu-id="4ed9f-152">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="4ed9f-153">String</span><span class="sxs-lookup"><span data-stu-id="4ed9f-153">String</span></span>|<span data-ttu-id="4ed9f-154">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="4ed9f-155">String</span><span class="sxs-lookup"><span data-stu-id="4ed9f-155">String</span></span>|<span data-ttu-id="4ed9f-156">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4ed9f-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ed9f-157">Requirements</span></span>

|<span data-ttu-id="4ed9f-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ed9f-158">Requirement</span></span>| <span data-ttu-id="4ed9f-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ed9f-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ed9f-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ed9f-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ed9f-161">1.3</span><span class="sxs-lookup"><span data-stu-id="4ed9f-161">1.3</span></span>|
|[<span data-ttu-id="4ed9f-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ed9f-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4ed9f-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ed9f-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4ed9f-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="4ed9f-164">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook17officeroamingsettings"></a><span data-ttu-id="4ed9f-165">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="4ed9f-165">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_7/office.RoamingSettings)</span></span>

<span data-ttu-id="4ed9f-166">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="4ed9f-167">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="4ed9f-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="4ed9f-168">Type</span><span class="sxs-lookup"><span data-stu-id="4ed9f-168">Type</span></span>

*   [<span data-ttu-id="4ed9f-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4ed9f-169">RoamingSettings</span></span>](/javascript/api/outlook_1_7/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="4ed9f-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="4ed9f-170">Requirements</span></span>

|<span data-ttu-id="4ed9f-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="4ed9f-171">Requirement</span></span>| <span data-ttu-id="4ed9f-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="4ed9f-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="4ed9f-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="4ed9f-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4ed9f-174">1.0</span><span class="sxs-lookup"><span data-stu-id="4ed9f-174">1.0</span></span>|
|[<span data-ttu-id="4ed9f-175">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="4ed9f-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4ed9f-176">Restreinte</span><span class="sxs-lookup"><span data-stu-id="4ed9f-176">Restricted</span></span>|
|[<span data-ttu-id="4ed9f-177">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="4ed9f-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="4ed9f-178">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="4ed9f-178">Compose or Read</span></span>|
