---
title: Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office
description: ''
ms.date: 11/26/2019
localization_priority: Priority
ms.openlocfilehash: fed23cfde1380e7c5728c78c995d3b89d44451f2
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629734"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="9ae7e-102">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="9ae7e-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="9ae7e-p101">Parfois, vos utilisateurs peuvent rencontrer des problèmes avec les compléments Office que vous développez. Par exemple, il se peut qu’un complément ne se charge pas ou soit inaccessible. Utilisez les informations de cet article pour résoudre les problèmes courants que vos utilisateurs rencontrent avec votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="9ae7e-106">Vous pouvez également utiliser [Fiddler](https://www.telerik.com/fiddler) pour identifier et déboguer les problèmes avec vos compléments.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="9ae7e-107">Erreurs courantes et étapes de dépannage</span><span class="sxs-lookup"><span data-stu-id="9ae7e-107">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="9ae7e-108">Le tableau suivant répertorie les messages d’erreur courants que les utilisateurs pourraient rencontrer, ainsi que les étapes que les utilisateurs peuvent suivre pour résoudre les erreurs.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-108">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="9ae7e-109">**Message d’erreur**</span><span class="sxs-lookup"><span data-stu-id="9ae7e-109">**Error message**</span></span>|<span data-ttu-id="9ae7e-110">**Solution**</span><span class="sxs-lookup"><span data-stu-id="9ae7e-110">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="9ae7e-111">Erreur d’application : impossible d’accéder au catalogue</span><span class="sxs-lookup"><span data-stu-id="9ae7e-111">App error: Catalog could not be reached</span></span>|<span data-ttu-id="9ae7e-p102">Vérifiez les paramètres de pare-feu. Le terme « catalogue » désigne AppSource. Ce message indique que l’utilisateur ne peut pas accéder à AppSource.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="9ae7e-p103">Erreur d’application : cette application n’a pas pu être démarrée. Fermez cette boîte de dialogue pour ignorer le problème, ou cliquez sur « Redémarrer » pour réessayer.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="9ae7e-116">Vérifiez que les dernières mises à jour d’Office sont installés, ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="9ae7e-116">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="9ae7e-117">Erreur : l’objet ne prend pas en charge la propriété ou la méthode « defineProperty »</span><span class="sxs-lookup"><span data-stu-id="9ae7e-117">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="9ae7e-p104">Vérifiez qu’Internet Explorer ne fonctionne pas en mode de compatibilité. Accédez à Outils >  **Paramètres d’affichage de compatibilité**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="9ae7e-p105">Désolé, nous n’avons pas pu charger l’application, car la version de votre navigateur n’est pas prise en charge. Cliquez ici pour obtenir la liste des versions de navigateur prises en charge.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="9ae7e-p106">Assurez-vous que le navigateur prend en charge le stockage local HTML5 ou réinitialisez les paramètres d’Internet Explorer. Pour plus d’informations sur les navigateurs pris en charge, reportez-vous à [Configuration requise pour exécuter des compléments Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|

## <a name="when-installing-an-add-in-you-see-error-loading-add-in-in-the-status-bar"></a><span data-ttu-id="9ae7e-124">Lors de l’installation d’un complément, le message « Erreur lors du chargement du complément » s’affiche dans la barre d’état</span><span class="sxs-lookup"><span data-stu-id="9ae7e-124">When installing an add-in, you see "Error loading add-in" in the status bar</span></span>

1. <span data-ttu-id="9ae7e-125">Fermez Office.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-125">Close Office.</span></span>
2. <span data-ttu-id="9ae7e-126">Vérifiez que le manifeste est valide.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-126">Verify that the manifest is valid</span></span>
3. <span data-ttu-id="9ae7e-127">Redémarrez le complément.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-127">Restart the add-in</span></span>
4. <span data-ttu-id="9ae7e-128">Réinstallez le complément.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-128">Install the add-in again.</span></span>

<span data-ttu-id="9ae7e-129">Vous pouvez également nous adresser des commentaires : si vous utilisez Excel sur Windows ou Mac, vous pouvez adresser un commentaire à l’équipe chargée de l’extensibilité d’Office directement à partir d’Excel.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-129">You can also give us feedback: if using Excel on Windows or Mac, you can report feedback to the Office extensibility team directly from Excel.</span></span> <span data-ttu-id="9ae7e-130">Pour ce faire, sélectionnez **Fichier** | **Commentaires** | **Envoyer un smiley mécontent**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-130">To do this, select **File** | **Feedback** | **Send a Frown**.</span></span> <span data-ttu-id="9ae7e-131">Envoyer un smiley mécontent fournit les journaux nécessaires pour comprendre le problème.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-131">Sending a frown provides the necessary logs to understand the issue.</span></span>

## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="9ae7e-132">Le complément Outlook ne fonctionne pas correctement</span><span class="sxs-lookup"><span data-stu-id="9ae7e-132">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="9ae7e-133">Si un complément Outlook s’exécutant sous Windows et [à l’aide d’Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) ne fonctionne pas correctement, essayez d’activer le débogage de script dans Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-133">If an Outlook add-in running on Windows and [using Internet Explorer](../concepts/browsers-used-by-office-web-add-ins.md) is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="9ae7e-134">Accédez à Outils >  **Options Internet** > **Avancées**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-134">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="9ae7e-135">Sous  **Parcourir**, décochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-135">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="9ae7e-p108">Nous vous recommandons de décocher ces paramètres uniquement pour résoudre le problème. Si vous ne les réactivez pas, vous recevrez des invites. Une fois que le problème est résolu, recochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p108">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="9ae7e-139">Le complément ne s’active pas dans Office 2013</span><span class="sxs-lookup"><span data-stu-id="9ae7e-139">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="9ae7e-140">Le complément ne s’active pas lorsque l’utilisateur effectue les étapes suivantes :</span><span class="sxs-lookup"><span data-stu-id="9ae7e-140">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="9ae7e-141">connexion à son compte Microsoft dans Office 2013 ;</span><span class="sxs-lookup"><span data-stu-id="9ae7e-141">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="9ae7e-142">activation de la vérification à deux étapes pour son compte Microsoft ;</span><span class="sxs-lookup"><span data-stu-id="9ae7e-142">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="9ae7e-143">vérification de son identité après invitation lorsqu’il tente d’insérer un complément.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-143">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="9ae7e-144">Pour résoudre ce problème, vérifiez que les dernières mises à jour Office sont installées ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="9ae7e-144">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="9ae7e-145">Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément</span><span class="sxs-lookup"><span data-stu-id="9ae7e-145">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="9ae7e-146">Consultez la rubrique relative à la [validation et à la résolution des problèmes de votre manifeste](troubleshoot-manifest.md) pour déboguer le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-146">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="9ae7e-147">La boîte de dialogue des compléments ne s’affiche pas</span><span class="sxs-lookup"><span data-stu-id="9ae7e-147">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="9ae7e-p109">Lorsqu’un utilisateur utilise un complément Office, il est invité à autoriser l’affichage d’une boîte de dialogue. L’utilisateur choisit **Autoriser** et le message d’erreur suivant apparaît :</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p109">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="9ae7e-p110">« Les paramètres de sécurité de votre navigateur nous empêchent de créer une boîte de dialogue. Essayez d’utiliser un autre navigateur, ou configurez votre navigateur de sorte que [URL] et le domaine affiché dans la barre d’adresse se trouvent dans la même zone de sécurité. »</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p110">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Capture d’écran du message d’erreur de la boîte de dialogue](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="9ae7e-153">**Navigateurs concernés**</span><span class="sxs-lookup"><span data-stu-id="9ae7e-153">**Affected browsers**</span></span>|<span data-ttu-id="9ae7e-154">**Plateformes concernées**</span><span class="sxs-lookup"><span data-stu-id="9ae7e-154">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="9ae7e-155">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="9ae7e-155">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="9ae7e-156">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="9ae7e-156">Office on the web</span></span>|

<span data-ttu-id="9ae7e-p111">Pour résoudre le problème, les utilisateurs finals et les administrateurs peuvent ajouter le domaine du complément à la liste des sites de confiance dans Internet Explorer. Utilisez la même procédure, que vous utilisiez le navigateur Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p111">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9ae7e-159">n’ajoutez pas l’URL d’un complément à votre liste de sites de confiance si vous ne faites pas confiance au complément.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-159">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="9ae7e-160">Pour ajouter une URL à votre liste de sites de confiance :</span><span class="sxs-lookup"><span data-stu-id="9ae7e-160">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="9ae7e-161">Dans \*\*Panneau de configuration, \*\*accédez à **Options Internet** > **Sécurité**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-161">In **Control Panel**, go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="9ae7e-162">Sélectionnez la zone **Sites de confiance**, puis choisissez **Sites**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-162">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="9ae7e-163">Entrez l’URL qui apparaît dans le message d’erreur, puis choisissez **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-163">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="9ae7e-p112">Essayez d’utiliser le complément à nouveau. Si le problème persiste, vérifiez les paramètres pour les autres zones de sécurité et assurez-vous que le domaine du complément se trouve dans la même zone que l’URL qui s’affiche dans la barre d’adresse de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p112">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="9ae7e-p113">Ce problème se produit lorsque l’API de la boîte de dialogue est utilisée en mode contextuel. Pour éviter ce problème, utilisez l’indicateur [displayInFrame](/javascript/api/office/office.ui). Cela nécessite que votre page prenne en charge l’affichage dans un iframe. L’exemple suivant montre comment utiliser l’indicateur.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p113">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="9ae7e-170">Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas</span><span class="sxs-lookup"><span data-stu-id="9ae7e-170">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="9ae7e-171">Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des éléments de menu, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-171">If changes you've made in the manifest, such as file names of ribbon button icons or text of menu items, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="9ae7e-172">Pour Windows :</span><span class="sxs-lookup"><span data-stu-id="9ae7e-172">For Windows:</span></span>
<span data-ttu-id="9ae7e-173">Supprimer le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-173">Delete the contents of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="9ae7e-174">Pour Mac :</span><span class="sxs-lookup"><span data-stu-id="9ae7e-174">For Mac:</span></span>

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="9ae7e-175">Pour iOS :</span><span class="sxs-lookup"><span data-stu-id="9ae7e-175">For iOS:</span></span>
<span data-ttu-id="9ae7e-p114">Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-p114">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="changes-to-static-files-such-as-javascript-html-and-css-do-not-take-effect"></a><span data-ttu-id="9ae7e-178">Les modifications apportées aux fichiers statiques, tels que JavaScript, HTML et CSS ne sont pas prises en compte.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-178">Changes to static files, such as JavaScript, HTML, and CSS do not take effect</span></span>

<span data-ttu-id="9ae7e-179">Le navigateur web met peut-être le contenu de ces fichiers en cache.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-179">The browser may be caching these files.</span></span> <span data-ttu-id="9ae7e-180">Pour éviter cela, vous pouvez désactiver la mise en cache côté client lors du développement.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-180">To prevent this, turn off client-side caching when developing.</span></span> <span data-ttu-id="9ae7e-181">Les spécifications dépendent du serveur utilisé.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-181">The details will depend on what kind of server you are using.</span></span> <span data-ttu-id="9ae7e-182">Dans la plupart des cas, elle implique l’ajout d’en-têtes aux réponses HTTP.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-182">In most cases, it involves adding certain headers to the HTTP Responses.</span></span> <span data-ttu-id="9ae7e-183">Nous vous recommandons d’exécuter les actions suivantes :</span><span class="sxs-lookup"><span data-stu-id="9ae7e-183">We suggest the following set:</span></span>

- <span data-ttu-id="9ae7e-184">Cache-Control : « privé, aucun cache, aucun magasin »</span><span class="sxs-lookup"><span data-stu-id="9ae7e-184">Cache-Control: "private, no-cache, no-store"</span></span>
- <span data-ttu-id="9ae7e-185">Pragma : « aucun cache »</span><span class="sxs-lookup"><span data-stu-id="9ae7e-185">Pragma: "no-cache"</span></span>
- <span data-ttu-id="9ae7e-186">Date d’expiration : « -1 »</span><span class="sxs-lookup"><span data-stu-id="9ae7e-186">Expires: "-1"</span></span>

<span data-ttu-id="9ae7e-187">Un exemple d’opération dans un serveur Node.JS Express est disponible dans [ce fichier app.js](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span><span class="sxs-lookup"><span data-stu-id="9ae7e-187">For an example of doing this in an Node.JS Express server, see [this app.js file](https://github.com/OfficeDev/Office-Add-in-NodeJS-SSO/blob/master/Complete/app.js).</span></span> <span data-ttu-id="9ae7e-188">Un exemple de projet ASP.NET est disponible dans [ce fichier cshtml](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/src/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span><span class="sxs-lookup"><span data-stu-id="9ae7e-188">For an example in an ASP.NET project, see [this cshtml file](https://github.com/OfficeDev/Office-Add-in-ASPNET-SSO/blob/master/src/Office-Add-in-ASPNET-SSO-WebAPI/Views/Shared/_Layout.cshtml).</span></span>

<span data-ttu-id="9ae7e-189">Si votre complément est hébergé dans Internet Information Server (IIS), vous pouvez également ajouter ce qui suit à web. config.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-189">If your add-in is hosted in Internet Information Server (IIS), you could also add the following to the web.config.</span></span>

```xml
<system.webServer>
  <staticContent>
    <clientCache cacheControlMode="UseMaxAge" cacheControlMaxAge="0.00:00:00" cacheControlCustom="must-revalidate" />
  </staticContent>
```

<span data-ttu-id="9ae7e-190">Si ces étapes ne semblent pas fonctionner au départ, vous devrez peut-être vider le cache du navigateur web.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-190">If these steps don't seem to work at first, you may need to clear the browser's cache.</span></span> <span data-ttu-id="9ae7e-191">Effectuez cette opération à l’aide de l’interface utilisateur du navigateur web.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-191">Do this through the UI of the browser.</span></span> <span data-ttu-id="9ae7e-192">Il est possible que le cache de périmètre ne soit pas correctement vidé lorsque vous essayez de le faire dans l’interface utilisateur Edge.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-192">Sometimes the Edge cache isn't successfully cleared when you try to clear it in the Edge UI.</span></span> <span data-ttu-id="9ae7e-193">Si cela se produit, exécutez la commande suivante dans l’invite de commandes Windows.</span><span class="sxs-lookup"><span data-stu-id="9ae7e-193">If that happens, run the following command in a Windows Command Prompt.</span></span>

```bash
del /s /f /q %LOCALAPPDATA%\Packages\Microsoft.Win32WebViewHost_cw5n1h2txyewy\AC\#!123\INetCache\
```

## <a name="see-also"></a><span data-ttu-id="9ae7e-194">Voir également</span><span class="sxs-lookup"><span data-stu-id="9ae7e-194">See also</span></span>

- [<span data-ttu-id="9ae7e-195">Débogage de compléments dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="9ae7e-195">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="9ae7e-196">Charger une version test d’un complément Office sur iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="9ae7e-196">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="9ae7e-197">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="9ae7e-197">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="9ae7e-198">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="9ae7e-198">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
