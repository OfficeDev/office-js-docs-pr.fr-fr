---
title: Office. Context-ensemble de conditions requises 1,6
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 35e2f69de7f94d96a1c2d4ae25ea482e892bb7fc
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064654"
---
# <a name="context"></a><span data-ttu-id="00d04-102">context</span><span class="sxs-lookup"><span data-stu-id="00d04-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="00d04-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="00d04-103">[Office](Office.md).context</span></span>

<span data-ttu-id="00d04-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="00d04-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="00d04-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00d04-106">Requirements</span></span>

|<span data-ttu-id="00d04-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00d04-107">Requirement</span></span>| <span data-ttu-id="00d04-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="00d04-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="00d04-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00d04-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00d04-110">1.0</span><span class="sxs-lookup"><span data-stu-id="00d04-110">1.0</span></span>|
|[<span data-ttu-id="00d04-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00d04-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00d04-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="00d04-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="00d04-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="00d04-113">Members and methods</span></span>

| <span data-ttu-id="00d04-114">Membre</span><span class="sxs-lookup"><span data-stu-id="00d04-114">Member</span></span> | <span data-ttu-id="00d04-115">Type</span><span class="sxs-lookup"><span data-stu-id="00d04-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="00d04-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="00d04-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="00d04-117">Member</span><span class="sxs-lookup"><span data-stu-id="00d04-117">Member</span></span> |
| [<span data-ttu-id="00d04-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="00d04-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="00d04-119">Membre</span><span class="sxs-lookup"><span data-stu-id="00d04-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="00d04-120">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="00d04-120">Namespaces</span></span>

<span data-ttu-id="00d04-121">[Mailbox](office.context.mailbox.md): permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="00d04-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="00d04-122">Members</span><span class="sxs-lookup"><span data-stu-id="00d04-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="00d04-123">displayLanguage: chaîne</span><span class="sxs-lookup"><span data-stu-id="00d04-123">displayLanguage: String</span></span>

<span data-ttu-id="00d04-124">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="00d04-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="00d04-125">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="00d04-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="00d04-126">Type</span><span class="sxs-lookup"><span data-stu-id="00d04-126">Type</span></span>

*   <span data-ttu-id="00d04-127">String</span><span class="sxs-lookup"><span data-stu-id="00d04-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="00d04-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00d04-128">Requirements</span></span>

|<span data-ttu-id="00d04-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00d04-129">Requirement</span></span>| <span data-ttu-id="00d04-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="00d04-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="00d04-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00d04-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00d04-132">1.0</span><span class="sxs-lookup"><span data-stu-id="00d04-132">1.0</span></span>|
|[<span data-ttu-id="00d04-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00d04-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00d04-134">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="00d04-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="00d04-135">Exemple</span><span class="sxs-lookup"><span data-stu-id="00d04-135">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-16"></a><span data-ttu-id="00d04-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="00d04-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.6)</span></span>

<span data-ttu-id="00d04-137">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="00d04-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="00d04-138">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="00d04-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="00d04-139">Type</span><span class="sxs-lookup"><span data-stu-id="00d04-139">Type</span></span>

*   [<span data-ttu-id="00d04-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="00d04-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="00d04-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="00d04-141">Requirements</span></span>

|<span data-ttu-id="00d04-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="00d04-142">Requirement</span></span>| <span data-ttu-id="00d04-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="00d04-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="00d04-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="00d04-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="00d04-145">1.0</span><span class="sxs-lookup"><span data-stu-id="00d04-145">1.0</span></span>|
|[<span data-ttu-id="00d04-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="00d04-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="00d04-147">Restreinte</span><span class="sxs-lookup"><span data-stu-id="00d04-147">Restricted</span></span>|
|[<span data-ttu-id="00d04-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="00d04-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="00d04-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="00d04-149">Compose or Read</span></span>|
