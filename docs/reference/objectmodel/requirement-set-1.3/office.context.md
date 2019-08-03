---
title: Office. Context-ensemble de conditions requises 1,3
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: c71be3f16ca637e9c1cb2790cda2a34956f6c67a
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064724"
---
# <a name="context"></a><span data-ttu-id="c6653-102">context</span><span class="sxs-lookup"><span data-stu-id="c6653-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="c6653-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="c6653-103">[Office](Office.md).context</span></span>

<span data-ttu-id="c6653-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="c6653-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6653-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c6653-106">Requirements</span></span>

|<span data-ttu-id="c6653-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c6653-107">Requirement</span></span>| <span data-ttu-id="c6653-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="c6653-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6653-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c6653-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6653-110">1.0</span><span class="sxs-lookup"><span data-stu-id="c6653-110">1.0</span></span>|
|[<span data-ttu-id="c6653-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c6653-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6653-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c6653-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="c6653-113">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="c6653-113">Namespaces</span></span>

<span data-ttu-id="c6653-114">[Mailbox](office.context.mailbox.md): permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="c6653-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="c6653-115">Members</span><span class="sxs-lookup"><span data-stu-id="c6653-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="c6653-116">displayLanguage: chaîne</span><span class="sxs-lookup"><span data-stu-id="c6653-116">displayLanguage: String</span></span>

<span data-ttu-id="c6653-117">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="c6653-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c6653-118">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="c6653-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c6653-119">Type</span><span class="sxs-lookup"><span data-stu-id="c6653-119">Type</span></span>

*   <span data-ttu-id="c6653-120">String</span><span class="sxs-lookup"><span data-stu-id="c6653-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c6653-121">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c6653-121">Requirements</span></span>

|<span data-ttu-id="c6653-122">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c6653-122">Requirement</span></span>| <span data-ttu-id="c6653-123">Valeur</span><span class="sxs-lookup"><span data-stu-id="c6653-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6653-124">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c6653-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6653-125">1.0</span><span class="sxs-lookup"><span data-stu-id="c6653-125">1.0</span></span>|
|[<span data-ttu-id="c6653-126">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c6653-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6653-127">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c6653-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c6653-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="c6653-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-13"></a><span data-ttu-id="c6653-129">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="c6653-129">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)</span></span>

<span data-ttu-id="c6653-130">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="c6653-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c6653-131">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="c6653-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c6653-132">Type</span><span class="sxs-lookup"><span data-stu-id="c6653-132">Type</span></span>

*   [<span data-ttu-id="c6653-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c6653-133">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="c6653-134">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="c6653-134">Requirements</span></span>

|<span data-ttu-id="c6653-135">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="c6653-135">Requirement</span></span>| <span data-ttu-id="c6653-136">Valeur</span><span class="sxs-lookup"><span data-stu-id="c6653-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="c6653-137">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="c6653-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c6653-138">1.0</span><span class="sxs-lookup"><span data-stu-id="c6653-138">1.0</span></span>|
|[<span data-ttu-id="c6653-139">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="c6653-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c6653-140">Restreinte</span><span class="sxs-lookup"><span data-stu-id="c6653-140">Restricted</span></span>|
|[<span data-ttu-id="c6653-141">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="c6653-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="c6653-142">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="c6653-142">Compose or Read</span></span>|
