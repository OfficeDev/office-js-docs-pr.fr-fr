---
title: Office.context-ensemble de conditions requises 1.1
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 392e54f1004bb395672c026ef749113f94ec7479
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432724"
---
# <a name="context"></a><span data-ttu-id="22b01-102">context</span><span class="sxs-lookup"><span data-stu-id="22b01-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="22b01-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="22b01-103">[Office](Office.md).context</span></span>

<span data-ttu-id="22b01-p101">L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API partagée](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="22b01-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="22b01-106">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22b01-106">Requirements</span></span>

|<span data-ttu-id="22b01-107">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22b01-107">Requirement</span></span>| <span data-ttu-id="22b01-108">Valeur</span><span class="sxs-lookup"><span data-stu-id="22b01-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="22b01-109">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22b01-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22b01-110">1.0</span><span class="sxs-lookup"><span data-stu-id="22b01-110">1.0</span></span>|
|[<span data-ttu-id="22b01-111">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22b01-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22b01-112">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22b01-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="22b01-113">Espaces de noms</span><span class="sxs-lookup"><span data-stu-id="22b01-113">Namespaces</span></span>

<span data-ttu-id="22b01-114">[mailbox](office.context.mailbox.md)- Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="22b01-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="22b01-115">Membres</span><span class="sxs-lookup"><span data-stu-id="22b01-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="22b01-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="22b01-116">displayLanguage :String</span></span>

<span data-ttu-id="22b01-117">Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="22b01-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="22b01-118">La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.</span><span class="sxs-lookup"><span data-stu-id="22b01-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="22b01-119">Type :</span><span class="sxs-lookup"><span data-stu-id="22b01-119">Type:</span></span>

*   <span data-ttu-id="22b01-120">Chaîne</span><span class="sxs-lookup"><span data-stu-id="22b01-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="22b01-121">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22b01-121">Requirements</span></span>

|<span data-ttu-id="22b01-122">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22b01-122">Requirement</span></span>| <span data-ttu-id="22b01-123">Valeur</span><span class="sxs-lookup"><span data-stu-id="22b01-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="22b01-124">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22b01-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22b01-125">1.0</span><span class="sxs-lookup"><span data-stu-id="22b01-125">1.0</span></span>|
|[<span data-ttu-id="22b01-126">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22b01-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22b01-127">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22b01-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="22b01-128">Exemple</span><span class="sxs-lookup"><span data-stu-id="22b01-128">Example</span></span>

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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a><span data-ttu-id="22b01-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="22b01-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span></span>

<span data-ttu-id="22b01-130">Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.</span><span class="sxs-lookup"><span data-stu-id="22b01-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="22b01-131">L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.</span><span class="sxs-lookup"><span data-stu-id="22b01-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="22b01-132">Type :</span><span class="sxs-lookup"><span data-stu-id="22b01-132">Type:</span></span>

*   [<span data-ttu-id="22b01-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="22b01-133">RoamingSettings</span></span>](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="22b01-134">Configuration requise</span><span class="sxs-lookup"><span data-stu-id="22b01-134">Requirements</span></span>

|<span data-ttu-id="22b01-135">Conditions requises</span><span class="sxs-lookup"><span data-stu-id="22b01-135">Requirement</span></span>| <span data-ttu-id="22b01-136">Valeur</span><span class="sxs-lookup"><span data-stu-id="22b01-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="22b01-137">Version de l’ensemble minimal de conditions de boîte aux lettres</span><span class="sxs-lookup"><span data-stu-id="22b01-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="22b01-138">1.0</span><span class="sxs-lookup"><span data-stu-id="22b01-138">1.0</span></span>|
|[<span data-ttu-id="22b01-139">Niveau d’autorisation minimal</span><span class="sxs-lookup"><span data-stu-id="22b01-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="22b01-140">Restreinte</span><span class="sxs-lookup"><span data-stu-id="22b01-140">Restricted</span></span>|
|[<span data-ttu-id="22b01-141">Mode Outlook applicable</span><span class="sxs-lookup"><span data-stu-id="22b01-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="22b01-142">Composition ou lecture</span><span class="sxs-lookup"><span data-stu-id="22b01-142">Compose or read</span></span>|