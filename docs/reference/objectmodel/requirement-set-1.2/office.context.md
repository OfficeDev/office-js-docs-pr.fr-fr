---
title: Office.context-ensemble de conditions requises 1.2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: cdf6b76ca090eb78d31fc1a6f328ce83c606e1f5
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35128425"
---
# <a name="context"></a><span data-ttu-id="48f31-102">context</span><span class="sxs-lookup"><span data-stu-id="48f31-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="48f31-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="48f31-103">[Office](Office.md).context</span></span>

<span data-ttu-id="48f31-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="48f31-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="48f31-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48f31-106">Requirements</span></span>

|<span data-ttu-id="48f31-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48f31-107">Requirement</span></span>| <span data-ttu-id="48f31-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="48f31-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f31-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48f31-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48f31-110">1.0</span><span class="sxs-lookup"><span data-stu-id="48f31-110">1.0</span></span>|
|[<span data-ttu-id="48f31-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48f31-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48f31-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="48f31-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="48f31-113">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="48f31-113">Namespaces</span></span>

<span data-ttu-id="48f31-114">[Mailbox](office.context.mailbox.md) -permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="48f31-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="48f31-115">Members</span><span class="sxs-lookup"><span data-stu-id="48f31-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="48f31-116">displayLanguage: chaîne</span><span class="sxs-lookup"><span data-stu-id="48f31-116">displayLanguage: String</span></span>

<span data-ttu-id="48f31-117">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="48f31-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="48f31-118">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="48f31-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="48f31-119">Type</span><span class="sxs-lookup"><span data-stu-id="48f31-119">Type</span></span>

*   <span data-ttu-id="48f31-120">String</span><span class="sxs-lookup"><span data-stu-id="48f31-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="48f31-121">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48f31-121">Requirements</span></span>

|<span data-ttu-id="48f31-122">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48f31-122">Requirement</span></span>| <span data-ttu-id="48f31-123">Valeur</span><span class="sxs-lookup"><span data-stu-id="48f31-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f31-124">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48f31-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48f31-125">1.0</span><span class="sxs-lookup"><span data-stu-id="48f31-125">1.0</span></span>|
|[<span data-ttu-id="48f31-126">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48f31-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48f31-127">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="48f31-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="48f31-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="48f31-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a><span data-ttu-id="48f31-129">roamingSettings: [roamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="48f31-129">roamingSettings: [RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span></span>

<span data-ttu-id="48f31-130">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="48f31-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="48f31-131">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="48f31-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="48f31-132">Type</span><span class="sxs-lookup"><span data-stu-id="48f31-132">Type</span></span>

*   [<span data-ttu-id="48f31-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="48f31-133">RoamingSettings</span></span>](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="48f31-134">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="48f31-134">Requirements</span></span>

|<span data-ttu-id="48f31-135">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="48f31-135">Requirement</span></span>| <span data-ttu-id="48f31-136">Valeur</span><span class="sxs-lookup"><span data-stu-id="48f31-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="48f31-137">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="48f31-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="48f31-138">1.0</span><span class="sxs-lookup"><span data-stu-id="48f31-138">1.0</span></span>|
|[<span data-ttu-id="48f31-139">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="48f31-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="48f31-140">Restreinte</span><span class="sxs-lookup"><span data-stu-id="48f31-140">Restricted</span></span>|
|[<span data-ttu-id="48f31-141">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="48f31-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="48f31-142">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="48f31-142">Compose or Read</span></span>|
