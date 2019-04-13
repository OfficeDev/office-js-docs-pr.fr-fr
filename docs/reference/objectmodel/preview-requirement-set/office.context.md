---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: a1e01142a4c0b84a4afcba89f76766d28595ba95
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838458"
---
# <a name="context"></a><span data-ttu-id="c3f66-102">context</span><span class="sxs-lookup"><span data-stu-id="c3f66-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="c3f66-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="c3f66-103">[Office](Office.md).context</span></span>

<span data-ttu-id="c3f66-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="c3f66-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3f66-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3f66-106">Requirements</span></span>

|<span data-ttu-id="c3f66-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3f66-107">Requirement</span></span>| <span data-ttu-id="c3f66-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3f66-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3f66-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3f66-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3f66-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c3f66-110">1.0</span></span>|
|[<span data-ttu-id="c3f66-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3f66-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3f66-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3f66-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c3f66-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="c3f66-113">Members and methods</span></span>

| <span data-ttu-id="c3f66-114">Membre</span><span class="sxs-lookup"><span data-stu-id="c3f66-114">Member</span></span> | <span data-ttu-id="c3f66-115">Type</span><span class="sxs-lookup"><span data-stu-id="c3f66-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c3f66-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c3f66-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c3f66-117">Member</span><span class="sxs-lookup"><span data-stu-id="c3f66-117">Member</span></span> |
| [<span data-ttu-id="c3f66-118">officeTheme</span><span class="sxs-lookup"><span data-stu-id="c3f66-118">officeTheme</span></span>](#officetheme-object) | <span data-ttu-id="c3f66-119">Member</span><span class="sxs-lookup"><span data-stu-id="c3f66-119">Member</span></span> |
| [<span data-ttu-id="c3f66-120">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c3f66-120">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c3f66-121">Membre</span><span class="sxs-lookup"><span data-stu-id="c3f66-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c3f66-122">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c3f66-122">Namespaces</span></span>

<span data-ttu-id="c3f66-123">[Mailbox](office.context.mailbox.md): permet d'accéder au modèle d'objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="c3f66-123">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="c3f66-124">Membres</span><span class="sxs-lookup"><span data-stu-id="c3f66-124">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="c3f66-125">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="c3f66-125">displayLanguage :String</span></span>

<span data-ttu-id="c3f66-126">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="c3f66-126">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c3f66-127">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="c3f66-127">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c3f66-128">Type</span><span class="sxs-lookup"><span data-stu-id="c3f66-128">Type</span></span>

*   <span data-ttu-id="c3f66-129">String</span><span class="sxs-lookup"><span data-stu-id="c3f66-129">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c3f66-130">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3f66-130">Requirements</span></span>

|<span data-ttu-id="c3f66-131">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3f66-131">Requirement</span></span>| <span data-ttu-id="c3f66-132">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3f66-132">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3f66-133">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3f66-133">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3f66-134">1.0</span><span class="sxs-lookup"><span data-stu-id="c3f66-134">1.0</span></span>|
|[<span data-ttu-id="c3f66-135">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3f66-135">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3f66-136">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3f66-136">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3f66-137">Exemple</span><span class="sxs-lookup"><span data-stu-id="c3f66-137">Example</span></span>

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

####  <a name="officetheme-object"></a><span data-ttu-id="c3f66-138">officeTheme :Object</span><span class="sxs-lookup"><span data-stu-id="c3f66-138">officeTheme :Object</span></span>

<span data-ttu-id="c3f66-139">Permet d’accéder aux propriétés pour les couleurs du thème Office.</span><span class="sxs-lookup"><span data-stu-id="c3f66-139">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="c3f66-140">Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.</span><span class="sxs-lookup"><span data-stu-id="c3f66-140">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="c3f66-p102">À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier > Compte Office > Thème Office**, qui est appliqué à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.</span><span class="sxs-lookup"><span data-stu-id="c3f66-p102">Using Office theme colors let's you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications. Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c3f66-143">Type</span><span class="sxs-lookup"><span data-stu-id="c3f66-143">Type</span></span>

*   <span data-ttu-id="c3f66-144">Objet</span><span class="sxs-lookup"><span data-stu-id="c3f66-144">Object</span></span>

##### <a name="properties"></a><span data-ttu-id="c3f66-145">Propriétés :</span><span class="sxs-lookup"><span data-stu-id="c3f66-145">Properties:</span></span>

|<span data-ttu-id="c3f66-146">Nom</span><span class="sxs-lookup"><span data-stu-id="c3f66-146">Name</span></span>| <span data-ttu-id="c3f66-147">Type</span><span class="sxs-lookup"><span data-stu-id="c3f66-147">Type</span></span>| <span data-ttu-id="c3f66-148">Description</span><span class="sxs-lookup"><span data-stu-id="c3f66-148">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="c3f66-149">String</span><span class="sxs-lookup"><span data-stu-id="c3f66-149">String</span></span>|<span data-ttu-id="c3f66-150">Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c3f66-150">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="c3f66-151">String</span><span class="sxs-lookup"><span data-stu-id="c3f66-151">String</span></span>|<span data-ttu-id="c3f66-152">Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c3f66-152">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="c3f66-153">String</span><span class="sxs-lookup"><span data-stu-id="c3f66-153">String</span></span>|<span data-ttu-id="c3f66-154">Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c3f66-154">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="c3f66-155">String</span><span class="sxs-lookup"><span data-stu-id="c3f66-155">String</span></span>|<span data-ttu-id="c3f66-156">Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.</span><span class="sxs-lookup"><span data-stu-id="c3f66-156">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c3f66-157">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3f66-157">Requirements</span></span>

|<span data-ttu-id="c3f66-158">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3f66-158">Requirement</span></span>| <span data-ttu-id="c3f66-159">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3f66-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3f66-160">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3f66-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3f66-161">1.3</span><span class="sxs-lookup"><span data-stu-id="c3f66-161">1.3</span></span>|
|[<span data-ttu-id="c3f66-162">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3f66-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3f66-163">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3f66-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c3f66-164">Exemple</span><span class="sxs-lookup"><span data-stu-id="c3f66-164">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="c3f66-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="c3f66-165">roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)</span></span>

<span data-ttu-id="c3f66-166">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c3f66-166">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c3f66-167">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="c3f66-167">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c3f66-168">Type</span><span class="sxs-lookup"><span data-stu-id="c3f66-168">Type</span></span>

*   [<span data-ttu-id="c3f66-169">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c3f66-169">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c3f66-170">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c3f66-170">Requirements</span></span>

|<span data-ttu-id="c3f66-171">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c3f66-171">Requirement</span></span>| <span data-ttu-id="c3f66-172">Valeur</span><span class="sxs-lookup"><span data-stu-id="c3f66-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="c3f66-173">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c3f66-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c3f66-174">1.0</span><span class="sxs-lookup"><span data-stu-id="c3f66-174">1.0</span></span>|
|[<span data-ttu-id="c3f66-175">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c3f66-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c3f66-176">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c3f66-176">Restricted</span></span>|
|[<span data-ttu-id="c3f66-177">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c3f66-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c3f66-178">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c3f66-178">Compose or Read</span></span>|
