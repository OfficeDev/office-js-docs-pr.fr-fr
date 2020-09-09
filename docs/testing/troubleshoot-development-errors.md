---
title: Résoudre les erreurs de développement avec les compléments Office
description: Découvrez comment résoudre les problèmes liés aux erreurs de développement dans les compléments Office.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 5801146165446352ec806f6f832e9976f96467ac
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409394"
---
# <a name="troubleshoot-development-errors-with-office-add-ins"></a><span data-ttu-id="9cee8-103">Résoudre les erreurs de développement avec les compléments Office</span><span class="sxs-lookup"><span data-stu-id="9cee8-103">Troubleshoot development errors with Office Add-ins</span></span>

## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="9cee8-104">Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément</span><span class="sxs-lookup"><span data-stu-id="9cee8-104">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="9cee8-105">Voir [Valider le manifeste d’un complément Office](troubleshoot-manifest.md) et [Déboguer votre complément avec la journalisation runtime](runtime-logging.md) pour déboguer les problèmes de manifeste de compléments.</span><span class="sxs-lookup"><span data-stu-id="9cee8-105">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="9cee8-106">Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas</span><span class="sxs-lookup"><span data-stu-id="9cee8-106">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="9cee8-107">Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des éléments de menu, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="9cee8-107">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="9cee8-108">Pour Windows :</span><span class="sxs-lookup"><span data-stu-id="9cee8-108">For Windows:</span></span>

<span data-ttu-id="9cee8-109">Supprimez le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` et supprimez le contenu du dossier `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\` , s’il existe.</span><span class="sxs-lookup"><span data-stu-id="9cee8-109">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`, and delete the contents of the folder `%userprofile%\AppData\Local\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\`, if it exists.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="9cee8-110">Pour Mac :</span><span class="sxs-lookup"><span data-stu-id="9cee8-110">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="9cee8-111">Pour iOS :</span><span class="sxs-lookup"><span data-stu-id="9cee8-111">For iOS:</span></span>
<span data-ttu-id="9cee8-p101">Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="9cee8-p101">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="9cee8-114">Les modifications apportées aux fichiers statiques, tels que JavaScript, HTML et CSS ne sont pas prises en compte.</span><span class="sxs-lookup"><span data-stu-id="9cee8-114">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="9cee8-115">Le navigateur web met peut-être le contenu de ces fichiers en cache.</span><span class="sxs-lookup"><span data-stu-id="9cee8-115">The browser may be caching these files.</span></span> <span data-ttu-id="9cee8-116">Pour éviter cela, vous pouvez désactiver la mise en cache côté client lors du développement.</span><span class="sxs-lookup"><span data-stu-id="9cee8-116">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="9cee8-117">Les spécifications dépendent du serveur utilisé.</span><span class="sxs-lookup"><span data-stu-id="9cee8-117">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="9cee8-118">Dans la plupart des cas, elle implique l’ajout d’en-têtes aux réponses HTTP.</span><span class="sxs-lookup"><span data-stu-id="9cee8-118">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="9cee8-119">Nous vous recommandons d’exécuter les actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="9cee8-119">We suggest the following set:</span></span>

- <span data-ttu-id="9cee8-120">Cache-Control : « privé, aucun cache, aucun magasin »</span><span class="sxs-lookup"><span data-stu-id="9cee8-120">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="9cee8-121">Pragma : « aucun cache »</span><span class="sxs-lookup"><span data-stu-id="9cee8-121">Pragma: "no-cache"</span></span>
- <span data-ttu-id="9cee8-122">Date d’expiration : « -1 »</span><span class="sxs-lookup"><span data-stu-id="9cee8-122">Expires: "-1"</span></span>

<span data-ttu-id="9cee8-123">Un exemple d’opération dans un serveur Node.JS Express est disponible dans [ce fichier app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span><span class="sxs-lookup"><span data-stu-id="9cee8-123">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="9cee8-124">Un exemple de projet ASP.NET est disponible dans [ce fichier cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span><span class="sxs-lookup"><span data-stu-id="9cee8-124">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="9cee8-125">Si votre complément est hébergé dans Internet Information Server (IIS), vous pouvez également ajouter ce qui suit à web. config.</span><span class="sxs-lookup"><span data-stu-id="9cee8-125">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="9cee8-126">Si ces étapes ne semblent pas fonctionner au départ, vous devrez peut-être vider le cache du navigateur web.</span><span class="sxs-lookup"><span data-stu-id="9cee8-126">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="9cee8-127">Effectuez cette opération à l’aide de l’interface utilisateur du navigateur web.</span><span class="sxs-lookup"><span data-stu-id="9cee8-127">Do this through the UI of the browser.</span></span> <span data-ttu-id="9cee8-128">Il est possible que le cache de périmètre ne soit pas correctement vidé lorsque vous essayez de le faire dans l’interface utilisateur Edge.</span><span class="sxs-lookup"><span data-stu-id="9cee8-128">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="9cee8-129">Si cela se produit, exécutez la commande suivante dans l’invite de commandes Windows.</span><span class="sxs-lookup"><span data-stu-id="9cee8-129">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="changes-made-to-property-values-dont-happen-and-there-is-no-error-message"></a><span data-ttu-id="9cee8-130">Les modifications apportées aux valeurs des propriétés ne se produisent pas et il n’y a pas de message d’erreur</span><span class="sxs-lookup"><span data-stu-id="9cee8-130">Changes made to property values don't happen and there is no error message</span></span>

<span data-ttu-id="9cee8-131">Consultez la documentation de référence pour savoir si la propriété est en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="9cee8-131">Check the reference documentation for the property to see if it is read only.</span></span> <span data-ttu-id="9cee8-132">En outre, les définitions de la [machine à écrire](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) pour Office js spécifient les propriétés d’objet en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="9cee8-132">Also, the [TypeScript definitions](../develop/referencing-the-javascript-api-for-office-library-from-its-cdn.md) for Office JS specify which object properties are read-only.</span></span> <span data-ttu-id="9cee8-133">Si vous tentez de définir une propriété en lecture seule, l’opération d’écriture échoue sans avertissement, sans qu’aucune erreur ne soit générée.</span><span class="sxs-lookup"><span data-stu-id="9cee8-133">If you attempt to set a read-only property, the write operation will fail silently, with no error thrown.</span></span> <span data-ttu-id="9cee8-134">L’exemple suivant tente à tort de définir la propriété en lecture seule [Chart.ID](/javascript/api/excel/excel.chart#id). Voir aussi [certaines propriétés ne peuvent pas être définies directement](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span><span class="sxs-lookup"><span data-stu-id="9cee8-134">The following example erroneously attempts to set the read-only property [Chart.id](/javascript/api/excel/excel.chart#id). See also [Some properties cannot be set directly](../develop/application-specific-api-model.md#some-properties-cannot-be-set-directly).</span></span>

```js
// This will do nothing, since `id` is a read-only property.
myChart.id = "5";
```

## <a name="add-in-doesnt-work-on-edge-but-it-works-on-other-browsers"></a><span data-ttu-id="9cee8-135">Le complément ne fonctionne pas sur Edge, mais fonctionne sur d’autres navigateurs</span><span class="sxs-lookup"><span data-stu-id="9cee8-135">Add-in doesn't work on Edge but it works on other browsers</span></span>

<span data-ttu-id="9cee8-136">Consultez la rubrique [Troubleshooting Microsoft Edge Problems](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span><span class="sxs-lookup"><span data-stu-id="9cee8-136">See [Troubleshooting Microsoft Edge issues](../concepts/browsers-used-by-office-web-add-ins.md#troubleshooting-microsoft-edge-issues).</span></span>

## <a name="excel-add-in-throws-errors-but-not-consistently"></a><span data-ttu-id="9cee8-137">Le complément Excel génère des erreurs, mais pas de façon cohérente</span><span class="sxs-lookup"><span data-stu-id="9cee8-137">Excel add-in throws errors, but not consistently</span></span>

<span data-ttu-id="9cee8-138">Consultez la rubrique [Troubleshoot Excel Add-ins](../excel/excel-add-ins-troubleshooting.md) pour obtenir les causes possibles.</span><span class="sxs-lookup"><span data-stu-id="9cee8-138">See [Troubleshoot Excel add-ins](../excel/excel-add-ins-troubleshooting.md) for possible causes.</span></span>

## <a name="see-also"></a><span data-ttu-id="9cee8-139">Voir également</span><span class="sxs-lookup"><span data-stu-id="9cee8-139">See also</span></span>

- [<span data-ttu-id="9cee8-140">Débogage de compléments dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="9cee8-140">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="9cee8-141">Charger une version test d’un complément Office sur iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="9cee8-141">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="9cee8-142">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="9cee8-142">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="9cee8-143">Complément Microsoft Office Extension de débogueur pour Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="9cee8-143">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="9cee8-144">Valider le manifeste d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="9cee8-144">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="9cee8-145">Déboguer votre complément avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="9cee8-145">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
- [<span data-ttu-id="9cee8-146">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="9cee8-146">Troubleshoot user errors with Office Add-ins</span></span>](testing-and-troubleshooting.md)
