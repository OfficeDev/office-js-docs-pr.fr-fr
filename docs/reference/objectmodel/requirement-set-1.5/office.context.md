---
title: Office. Context-ensemble de conditions requises 1,5
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: ee795787b6c42ff331161a4641c7aa35ae960368
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696084"
---
# <a name="context"></a><span data-ttu-id="b8190-102">context</span><span class="sxs-lookup"><span data-stu-id="b8190-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="b8190-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="b8190-103">[Office](Office.md).context</span></span>

<span data-ttu-id="b8190-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="b8190-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8190-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b8190-106">Requirements</span></span>

|<span data-ttu-id="b8190-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b8190-107">Requirement</span></span>| <span data-ttu-id="b8190-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="b8190-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8190-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b8190-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8190-110">1.0</span><span class="sxs-lookup"><span data-stu-id="b8190-110">1.0</span></span>|
|[<span data-ttu-id="b8190-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b8190-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8190-112">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b8190-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b8190-113">Membres et méthodes</span><span class="sxs-lookup"><span data-stu-id="b8190-113">Members and methods</span></span>

| <span data-ttu-id="b8190-114">Membre</span><span class="sxs-lookup"><span data-stu-id="b8190-114">Member</span></span> | <span data-ttu-id="b8190-115">Type</span><span class="sxs-lookup"><span data-stu-id="b8190-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b8190-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b8190-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b8190-117">Member</span><span class="sxs-lookup"><span data-stu-id="b8190-117">Member</span></span> |
| [<span data-ttu-id="b8190-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b8190-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b8190-119">Membre</span><span class="sxs-lookup"><span data-stu-id="b8190-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b8190-120">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="b8190-120">Namespaces</span></span>

<span data-ttu-id="b8190-121">[Mailbox](office.context.mailbox.md): permet d’accéder au modèle d’objet du complément Outlook pour Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="b8190-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="b8190-122">Members</span><span class="sxs-lookup"><span data-stu-id="b8190-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="b8190-123">displayLanguage: chaîne</span><span class="sxs-lookup"><span data-stu-id="b8190-123">displayLanguage: String</span></span>

<span data-ttu-id="b8190-124">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="b8190-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="b8190-125">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="b8190-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="b8190-126">Type</span><span class="sxs-lookup"><span data-stu-id="b8190-126">Type</span></span>

*   <span data-ttu-id="b8190-127">String</span><span class="sxs-lookup"><span data-stu-id="b8190-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b8190-128">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b8190-128">Requirements</span></span>

|<span data-ttu-id="b8190-129">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b8190-129">Requirement</span></span>| <span data-ttu-id="b8190-130">Valeur</span><span class="sxs-lookup"><span data-stu-id="b8190-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8190-131">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b8190-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8190-132">1.0</span><span class="sxs-lookup"><span data-stu-id="b8190-132">1.0</span></span>|
|[<span data-ttu-id="b8190-133">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b8190-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8190-134">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b8190-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b8190-135">Exemple</span><span class="sxs-lookup"><span data-stu-id="b8190-135">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-15"></a><span data-ttu-id="b8190-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="b8190-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.5)</span></span>

<span data-ttu-id="b8190-137">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="b8190-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b8190-138">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="b8190-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b8190-139">Type</span><span class="sxs-lookup"><span data-stu-id="b8190-139">Type</span></span>

*   [<span data-ttu-id="b8190-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b8190-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.5)

##### <a name="requirements"></a><span data-ttu-id="b8190-141">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="b8190-141">Requirements</span></span>

|<span data-ttu-id="b8190-142">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="b8190-142">Requirement</span></span>| <span data-ttu-id="b8190-143">Valeur</span><span class="sxs-lookup"><span data-stu-id="b8190-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="b8190-144">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="b8190-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b8190-145">1.0</span><span class="sxs-lookup"><span data-stu-id="b8190-145">1.0</span></span>|
|[<span data-ttu-id="b8190-146">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="b8190-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b8190-147">Restreinte</span><span class="sxs-lookup"><span data-stu-id="b8190-147">Restricted</span></span>|
|[<span data-ttu-id="b8190-148">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="b8190-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b8190-149">Rédaction ou lecture</span><span class="sxs-lookup"><span data-stu-id="b8190-149">Compose or Read</span></span>|
