---
title: Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office
description: Découvrez comment résoudre les problèmes liés aux erreurs utilisateur dans les compléments Office.
ms.date: 06/17/2020
localization_priority: Normal
ms.openlocfilehash: 51f5ec406a09b18ece24b74dc22718e7fd422e38
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159184"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="a57a8-103">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="a57a8-103">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="a57a8-p101">Parfois, vos utilisateurs peuvent rencontrer des problèmes avec les compléments Office que vous développez. Par exemple, il se peut qu’un complément ne se charge pas ou soit inaccessible. Utilisez les informations de cet article pour résoudre les problèmes courants que vos utilisateurs rencontrent avec votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="a57a8-107">Vous pouvez également utiliser [Fiddler](https://www.telerik.com/fiddler) pour identifier et déboguer les problèmes avec vos compléments.</span><span class="sxs-lookup"><span data-stu-id="a57a8-107">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="a57a8-108">Erreurs courantes et étapes de dépannage</span><span class="sxs-lookup"><span data-stu-id="a57a8-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="a57a8-109">Le tableau suivant répertorie les messages d’erreur courants que les utilisateurs pourraient rencontrer, ainsi que les étapes que les utilisateurs peuvent suivre pour résoudre les erreurs.</span><span class="sxs-lookup"><span data-stu-id="a57a8-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="a57a8-110">**Message d’erreur**</span><span class="sxs-lookup"><span data-stu-id="a57a8-110">**Error message**</span></span>|<span data-ttu-id="a57a8-111">**Solution**</span><span class="sxs-lookup"><span data-stu-id="a57a8-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="a57a8-112">Erreur d’application : impossible d’accéder au catalogue</span><span class="sxs-lookup"><span data-stu-id="a57a8-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="a57a8-p102">Vérifiez les paramètres de pare-feu. Le terme « catalogue » désigne AppSource. Ce message indique que l’utilisateur ne peut pas accéder à AppSource.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="a57a8-p103">Erreur d’application : cette application n’a pas pu être démarrée. Fermez cette boîte de dialogue pour ignorer le problème, ou cliquez sur « Redémarrer » pour réessayer.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="a57a8-117">Vérifiez que les dernières mises à jour d’Office sont installés, ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="a57a8-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="a57a8-118">Erreur : l’objet ne prend pas en charge la propriété ou la méthode « defineProperty »</span><span class="sxs-lookup"><span data-stu-id="a57a8-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="a57a8-119">Vérifiez qu’Internet Explorer ne fonctionne pas en mode de compatibilité.</span><span class="sxs-lookup"><span data-stu-id="a57a8-119">Confirm that Internet Explorer is not running in Compatibility Mode.</span></span> <span data-ttu-id="a57a8-120">Accédez à Outils > **Paramètres d’affichage de compatibilité**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-120">Go to Tools > **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="a57a8-p105">Désolé, nous n’avons pas pu charger l’application, car la version de votre navigateur n’est pas prise en charge. Cliquez ici pour obtenir la liste des versions de navigateur prises en charge.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="a57a8-p106">Assurez-vous que le navigateur prend en charge le stockage local HTML5 ou réinitialisez les paramètres d’Internet Explorer. Pour plus d’informations sur les navigateurs pris en charge, reportez-vous à [Configuration requise pour exécuter des compléments Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="a57a8-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="a57a8-125">Lors de l’installation d’un complément, le message « Erreur lors du chargement du complément » s’affiche dans la barre d’état</span><span class="sxs-lookup"><span data-stu-id="a57a8-125">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="a57a8-126">Fermez Office.</span><span class="sxs-lookup"><span data-stu-id="a57a8-126">Close Office.</span></span>
2. <span data-ttu-id="a57a8-127">Vérifiez que le manifeste est valide.</span><span class="sxs-lookup"><span data-stu-id="a57a8-127">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="a57a8-128">Redémarrez le complément.</span><span class="sxs-lookup"><span data-stu-id="a57a8-128">Restart the add-in</span></span>
4. <span data-ttu-id="a57a8-129">Réinstallez le complément.</span><span class="sxs-lookup"><span data-stu-id="a57a8-129">Install the add-in again.</span></span>

<span data-ttu-id="a57a8-130">Vous pouvez également nous adresser des commentaires : si vous utilisez Excel sur Windows ou Mac, vous pouvez adresser un commentaire à l’équipe chargée de l’extensibilité d’Office directement à partir d’Excel.</span><span class="sxs-lookup"><span data-stu-id="a57a8-130">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="a57a8-131">Pour ce faire, sélectionnez **Fichier** | **Commentaires** | **Envoyer un smiley mécontent**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-131">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="a57a8-132">Envoyer un smiley mécontent fournit les journaux nécessaires pour comprendre le problème.</span><span class="sxs-lookup"><span data-stu-id="a57a8-132">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="a57a8-133">Le complément Outlook ne fonctionne pas correctement</span><span class="sxs-lookup"><span data-stu-id="a57a8-133">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="a57a8-134">Si un complément Outlook s’exécutant sous Windows et [à l’aide d’Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) ne fonctionne pas correctement, essayez d’activer le débogage de script dans Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="a57a8-134">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="a57a8-135">Accédez à outils > **Internet options**  >  **avancées**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-135">Go to Tools > **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="a57a8-136">Sous **Parcourir**, décochez les cases **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-136">Under **Browsing**, uncheck **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="a57a8-p108">Nous vous recommandons de décocher ces paramètres uniquement pour résoudre le problème. Si vous ne les réactivez pas, vous recevrez des invites. Une fois que le problème est résolu, recochez les cases **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="a57a8-140">Le complément ne s’active pas dans Office 2013</span><span class="sxs-lookup"><span data-stu-id="a57a8-140">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="a57a8-141">Le complément ne s’active pas lorsque l’utilisateur effectue les étapes suivantes :</span><span class="sxs-lookup"><span data-stu-id="a57a8-141">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="a57a8-142">connexion à son compte Microsoft dans Office 2013 ;</span><span class="sxs-lookup"><span data-stu-id="a57a8-142">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="a57a8-143">activation de la vérification à deux étapes pour son compte Microsoft ;</span><span class="sxs-lookup"><span data-stu-id="a57a8-143">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="a57a8-144">vérification de son identité après invitation lorsqu’il tente d’insérer un complément.</span><span class="sxs-lookup"><span data-stu-id="a57a8-144">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="a57a8-145">Pour résoudre ce problème, vérifiez que les dernières mises à jour Office sont installées ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="a57a8-145">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="a57a8-146">Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément</span><span class="sxs-lookup"><span data-stu-id="a57a8-146">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="a57a8-147">Voir [Valider le manifeste d’un complément Office](troubleshoot-manifest.md) et [Déboguer votre complément avec la journalisation runtime](runtime-logging.md) pour déboguer les problèmes de manifeste de compléments.</span><span class="sxs-lookup"><span data-stu-id="a57a8-147">See [Validate an Office Add-in's manifest](troubleshoot-manifest.md) and [Debug your add-in with runtime logging](runtime-logging.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="a57a8-148">La boîte de dialogue des compléments ne s’affiche pas</span><span class="sxs-lookup"><span data-stu-id="a57a8-148">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="a57a8-p109">Lorsqu’un utilisateur utilise un complément Office, il est invité à autoriser l’affichage d’une boîte de dialogue. L’utilisateur choisit **Autoriser** et le message d’erreur suivant apparaît :</span><span class="sxs-lookup"><span data-stu-id="a57a8-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="a57a8-p110">« Les paramètres de sécurité de votre navigateur nous empêchent de créer une boîte de dialogue. Essayez d’utiliser un autre navigateur, ou configurez votre navigateur de sorte que [URL] et le domaine affiché dans la barre d’adresse se trouvent dans la même zone de sécurité. »</span><span class="sxs-lookup"><span data-stu-id="a57a8-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Capture d’écran du message d’erreur de la boîte de dialogue](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="a57a8-154">**Navigateurs concernés**</span><span class="sxs-lookup"><span data-stu-id="a57a8-154">**Affected browsers**</span></span>|<span data-ttu-id="a57a8-155">**Plateformes concernées**</span><span class="sxs-lookup"><span data-stu-id="a57a8-155">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="a57a8-156">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="a57a8-156">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="a57a8-157">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="a57a8-157">Office on the web</span></span>|

<span data-ttu-id="a57a8-p111">Pour résoudre le problème, les utilisateurs finals et les administrateurs peuvent ajouter le domaine du complément à la liste des sites de confiance dans Internet Explorer. Utilisez la même procédure, que vous utilisiez le navigateur Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="a57a8-160">n’ajoutez pas l’URL d’un complément à votre liste de sites de confiance si vous ne faites pas confiance au complément.</span><span class="sxs-lookup"><span data-stu-id="a57a8-160">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="a57a8-161">Pour ajouter une URL à votre liste de sites de confiance :</span><span class="sxs-lookup"><span data-stu-id="a57a8-161">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="a57a8-162">Dans \*\*Panneau de configuration, \*\*accédez à **Options Internet** > **Sécurité**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-162">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="a57a8-163">Sélectionnez la zone **Sites de confiance**, puis choisissez **Sites**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-163">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="a57a8-164">Entrez l’URL qui apparaît dans le message d’erreur, puis choisissez **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="a57a8-164">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="a57a8-p112">Essayez d’utiliser le complément à nouveau. Si le problème persiste, vérifiez les paramètres pour les autres zones de sécurité et assurez-vous que le domaine du complément se trouve dans la même zone que l’URL qui s’affiche dans la barre d’adresse de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="a57a8-p113">Ce problème se produit lorsque l’API de la boîte de dialogue est utilisée en mode contextuel. Pour éviter ce problème, utilisez l’indicateur [displayInFrame](/javascript/api/office/office.ui). Cela nécessite que votre page prenne en charge l’affichage dans un iframe. L’exemple suivant montre comment utiliser l’indicateur.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInIFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="a57a8-171">Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas</span><span class="sxs-lookup"><span data-stu-id="a57a8-171">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="a57a8-172">Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des éléments de menu, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="a57a8-172">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="a57a8-173">Pour Windows :</span><span class="sxs-lookup"><span data-stu-id="a57a8-173">For Windows:</span></span>
<span data-ttu-id="a57a8-174">Supprimer le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="a57a8-174">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="a57a8-175">Pour Mac :</span><span class="sxs-lookup"><span data-stu-id="a57a8-175">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="a57a8-176">Pour iOS :</span><span class="sxs-lookup"><span data-stu-id="a57a8-176">For iOS:</span></span>
<span data-ttu-id="a57a8-p114">Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="a57a8-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="a57a8-179">Les modifications apportées aux fichiers statiques, tels que JavaScript, HTML et CSS ne sont pas prises en compte.</span><span class="sxs-lookup"><span data-stu-id="a57a8-179">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="a57a8-180">Le navigateur web met peut-être le contenu de ces fichiers en cache.</span><span class="sxs-lookup"><span data-stu-id="a57a8-180">The browser may be caching these files.</span></span> <span data-ttu-id="a57a8-181">Pour éviter cela, vous pouvez désactiver la mise en cache côté client lors du développement.</span><span class="sxs-lookup"><span data-stu-id="a57a8-181">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="a57a8-182">Les spécifications dépendent du serveur utilisé.</span><span class="sxs-lookup"><span data-stu-id="a57a8-182">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="a57a8-183">Dans la plupart des cas, elle implique l’ajout d’en-têtes aux réponses HTTP.</span><span class="sxs-lookup"><span data-stu-id="a57a8-183">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="a57a8-184">Nous vous recommandons d’exécuter les actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="a57a8-184">We suggest the following set:</span></span>

- <span data-ttu-id="a57a8-185">Cache-Control : « privé, aucun cache, aucun magasin »</span><span class="sxs-lookup"><span data-stu-id="a57a8-185">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="a57a8-186">Pragma : « aucun cache »</span><span class="sxs-lookup"><span data-stu-id="a57a8-186">Pragma: "no-cache"</span></span>
- <span data-ttu-id="a57a8-187">Date d’expiration : « -1 »</span><span class="sxs-lookup"><span data-stu-id="a57a8-187">Expires: "-1"</span></span>

<span data-ttu-id="a57a8-188">Un exemple d’opération dans un serveur Node.JS Express est disponible dans [ce fichier app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span><span class="sxs-lookup"><span data-stu-id="a57a8-188">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="a57a8-189">Un exemple de projet ASP.NET est disponible dans [ce fichier cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span><span class="sxs-lookup"><span data-stu-id="a57a8-189">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/Complete/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="a57a8-190">Si votre complément est hébergé dans Internet Information Server (IIS), vous pouvez également ajouter ce qui suit à web. config.</span><span class="sxs-lookup"><span data-stu-id="a57a8-190">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="a57a8-191">Si ces étapes ne semblent pas fonctionner au départ, vous devrez peut-être vider le cache du navigateur web.</span><span class="sxs-lookup"><span data-stu-id="a57a8-191">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="a57a8-192">Effectuez cette opération à l’aide de l’interface utilisateur du navigateur web.</span><span class="sxs-lookup"><span data-stu-id="a57a8-192">Do this through the UI of the browser.</span></span> <span data-ttu-id="a57a8-193">Il est possible que le cache de périmètre ne soit pas correctement vidé lorsque vous essayez de le faire dans l’interface utilisateur Edge.</span><span class="sxs-lookup"><span data-stu-id="a57a8-193">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="a57a8-194">Si cela se produit, exécutez la commande suivante dans l’invite de commandes Windows.</span><span class="sxs-lookup"><span data-stu-id="a57a8-194">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a><span data-ttu-id="a57a8-195">Voir également</span><span class="sxs-lookup"><span data-stu-id="a57a8-195">See also</span></span>

- [<span data-ttu-id="a57a8-196">Débogage de compléments dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="a57a8-196">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md)
- [<span data-ttu-id="a57a8-197">Charger une version test d’un complément Office sur iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="a57a8-197">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="a57a8-198">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="a57a8-198">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="a57a8-199">Extension du débogueur de complément Microsoft Office pour Visual Studio code</span><span class="sxs-lookup"><span data-stu-id="a57a8-199">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
- [<span data-ttu-id="a57a8-200">Valider le manifeste d’un complément Office</span><span class="sxs-lookup"><span data-stu-id="a57a8-200">Validate an Office Add-in's manifest</span></span>](troubleshoot-manifest.md)
- [<span data-ttu-id="a57a8-201">Déboguer votre complément avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="a57a8-201">Debug your add-in with runtime logging</span></span>](runtime-logging.md)
