---
title: Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office
description: ''
ms.date: 06/20/2019
localization_priority: Priority
ms.openlocfilehash: a82dc05789b4c35a954337a64197d3ac1a190b96
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35126903"
---
# <a name="troubleshoot-user-errors-with-office-add-ins"></a><span data-ttu-id="25213-102">Résolution des erreurs rencontrées par l’utilisateur avec des compléments Office</span><span class="sxs-lookup"><span data-stu-id="25213-102">Troubleshoot user errors with Office Add-ins</span></span>

<span data-ttu-id="25213-p101">Parfois, vos utilisateurs peuvent rencontrer des problèmes avec les compléments Office que vous développez. Par exemple, il se peut qu’un complément ne se charge pas ou soit inaccessible. Utilisez les informations de cet article pour résoudre les problèmes courants que vos utilisateurs rencontrent avec votre complément Office.</span><span class="sxs-lookup"><span data-stu-id="25213-p101">At times your users might encounter issues with Office Add-ins that you develop. For example, an add-in fails to load or is inaccessible. Use the information in this article to help resolve common issues that your users encounter with your Office Add-in.</span></span> 

<span data-ttu-id="25213-106">Vous pouvez également utiliser [Fiddler](https://www.telerik.com/fiddler) pour identifier et déboguer les problèmes avec vos compléments.</span><span class="sxs-lookup"><span data-stu-id="25213-106">You can also use [Fiddler](https://www.telerik.com/fiddler) to identify and debug issues with your add-ins.</span></span>

<span data-ttu-id="25213-107">Une fois le problème de l’utilisateur résolu, vous pouvez [répondre directement aux avis des clients dans AppSource](/office/dev/store/create-effective-office-store-listings).</span><span class="sxs-lookup"><span data-stu-id="25213-107">After you resolve the user's issue, you can [respond directly to customer reviews in AppSource](/office/dev/store/create-effective-office-store-listings).</span></span>

## <a name="common-errors-and-troubleshooting-steps"></a><span data-ttu-id="25213-108">Erreurs courantes et étapes de dépannage</span><span class="sxs-lookup"><span data-stu-id="25213-108">Common errors and troubleshooting steps</span></span>

<span data-ttu-id="25213-109">Le tableau suivant répertorie les messages d’erreur courants que les utilisateurs pourraient rencontrer, ainsi que les étapes que les utilisateurs peuvent suivre pour résoudre les erreurs.</span><span class="sxs-lookup"><span data-stu-id="25213-109">The following table lists common error messages that users might encounter and steps that your users can take to resolve the errors.</span></span>



|<span data-ttu-id="25213-110">**Message d’erreur**</span><span class="sxs-lookup"><span data-stu-id="25213-110">**Error message**</span></span>|<span data-ttu-id="25213-111">**Solution**</span><span class="sxs-lookup"><span data-stu-id="25213-111">**Resolution**</span></span>|
|:-----|:-----|
|<span data-ttu-id="25213-112">Erreur d’application : impossible d’accéder au catalogue</span><span class="sxs-lookup"><span data-stu-id="25213-112">App error: Catalog could not be reached</span></span>|<span data-ttu-id="25213-p102">Vérifiez les paramètres de pare-feu. Le terme « catalogue » désigne AppSource. Ce message indique que l’utilisateur ne peut pas accéder à AppSource.</span><span class="sxs-lookup"><span data-stu-id="25213-p102">Verify firewall settings."Catalog" refers to AppSource. This message indicates that the user cannot access AppSource.</span></span>|
|<span data-ttu-id="25213-p103">Erreur d’application : cette application n’a pas pu être démarrée. Fermez cette boîte de dialogue pour ignorer le problème, ou cliquez sur « Redémarrer » pour réessayer.</span><span class="sxs-lookup"><span data-stu-id="25213-p103">APP ERROR: This app could not be started. Close this dialog to ignore the problem or click "Restart" to try again.</span></span>|<span data-ttu-id="25213-117">Vérifiez que les dernières mises à jour d’Office sont installés, ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="25213-117">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>|
|<span data-ttu-id="25213-118">Erreur : l’objet ne prend pas en charge la propriété ou la méthode « defineProperty »</span><span class="sxs-lookup"><span data-stu-id="25213-118">Error: Object doesn't support property or method 'defineProperty'</span></span>|<span data-ttu-id="25213-p104">Vérifiez qu’Internet Explorer ne fonctionne pas en mode de compatibilité. Accédez à Outils >  **Paramètres d’affichage de compatibilité**.</span><span class="sxs-lookup"><span data-stu-id="25213-p104">Confirm that Internet Explorer is not running in Compatibility Mode. Go to Tools >  **Compatibility View Settings**.</span></span>|
|<span data-ttu-id="25213-p105">Désolé, nous n’avons pas pu charger l’application, car la version de votre navigateur n’est pas prise en charge. Cliquez ici pour obtenir la liste des versions de navigateur prises en charge.</span><span class="sxs-lookup"><span data-stu-id="25213-p105">Sorry, we couldn't load the app because your browser version is not supported. Click here for a list of supported browser versions.</span></span>|<span data-ttu-id="25213-p106">Assurez-vous que le navigateur prend en charge le stockage local HTML5 ou réinitialisez les paramètres d’Internet Explorer. Pour plus d’informations sur les navigateurs pris en charge, reportez-vous à [Configuration requise pour exécuter des compléments Office](../concepts/requirements-for-running-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="25213-p106">Make sure that the browser supports HTML5 local storage, or reset your Internet Explorer settings. For information about supported browsers, see [Requirements for running Office Add-ins](../concepts/requirements-for-running-office-add-ins.md).</span></span>|


## <a name="outlook-add-in-doesnt-work-correctly"></a><span data-ttu-id="25213-125">§LTA Le complément Outlook ne fonctionne pas correctement</span><span class="sxs-lookup"><span data-stu-id="25213-125">Outlook add-in doesn't work correctly</span></span>

<span data-ttu-id="25213-126">Si un complément Outlook s’exécutant sous Windows ne fonctionne pas correctement, essayez d’activer le débogage de script dans Internet Explorer.</span><span class="sxs-lookup"><span data-stu-id="25213-126">If an Outlook add-in running on Windows is not working correctly, try turning on script debugging in Internet Explorer.</span></span> 


- <span data-ttu-id="25213-127">Accédez à Outils >  **Options Internet** > **Avancées**.</span><span class="sxs-lookup"><span data-stu-id="25213-127">Go to Tools >  **Internet Options** > **Advanced**.</span></span>
    
- <span data-ttu-id="25213-128">Sous  **Parcourir**, décochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.</span><span class="sxs-lookup"><span data-stu-id="25213-128">Under  **Browsing**, uncheck  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)**.</span></span>
    
<span data-ttu-id="25213-p107">Nous vous recommandons de décocher ces paramètres uniquement pour résoudre le problème. Si vous ne les réactivez pas, vous recevrez des invites. Une fois que le problème est résolu, recochez les cases  **Désactiver le débogage des scripts (Internet Explorer)** et **Désactiver le débogage des scripts (autres applications)**.</span><span class="sxs-lookup"><span data-stu-id="25213-p107">We recommend that you uncheck these settings only to troubleshoot the issue. If you leave them unchecked, you will get prompts when you browse. After the issue is resolved, check  **Disable script debugging (Internet Explorer)** and **Disable script debugging (Other)** again.</span></span>


## <a name="add-in-doesnt-activate-in-office-2013"></a><span data-ttu-id="25213-132">Le complément ne s’active pas dans Office 2013</span><span class="sxs-lookup"><span data-stu-id="25213-132">Add-in doesn't activate in Office 2013</span></span>

<span data-ttu-id="25213-133">Le complément ne s’active pas lorsque l’utilisateur effectue les étapes suivantes :</span><span class="sxs-lookup"><span data-stu-id="25213-133">If the add-in doesn't activate when the user performs the following steps:</span></span>


1. <span data-ttu-id="25213-134">connexion à son compte Microsoft dans Office 2013 ;</span><span class="sxs-lookup"><span data-stu-id="25213-134">Signs in with their Microsoft account in Office 2013.</span></span>
    
2. <span data-ttu-id="25213-135">activation de la vérification à deux étapes pour son compte Microsoft ;</span><span class="sxs-lookup"><span data-stu-id="25213-135">Enables two-step verification for their Microsoft account.</span></span>
    
3. <span data-ttu-id="25213-136">vérification de son identité après invitation lorsqu’il tente d’insérer un complément.</span><span class="sxs-lookup"><span data-stu-id="25213-136">Verifies their identity when prompted when they try to insert an add-in.</span></span>
    
<span data-ttu-id="25213-137">Pour résoudre ce problème, vérifiez que les dernières mises à jour Office sont installées ou téléchargez la [mise à jour pour Office 2013](https://support.microsoft.com/kb/2986156/).</span><span class="sxs-lookup"><span data-stu-id="25213-137">Verify that the latest Office updates are installed, or download the [update for Office 2013](https://support.microsoft.com/kb/2986156/).</span></span>


## <a name="add-in-doesnt-load-in-task-pane-or-other-issues-with-the-add-in-manifest"></a><span data-ttu-id="25213-138">Le complément ne se charge pas dans le volet des tâches ou d’autres problèmes existent avec le manifeste du complément</span><span class="sxs-lookup"><span data-stu-id="25213-138">Add-in doesn't load in task pane or other issues with the add-in manifest</span></span>

<span data-ttu-id="25213-139">Consultez la rubrique relative à la [validation et à la résolution des problèmes de votre manifeste](troubleshoot-manifest.md) pour déboguer le manifeste de votre complément.</span><span class="sxs-lookup"><span data-stu-id="25213-139">See [Validate and troubleshoot issues with your manifest](troubleshoot-manifest.md) to debug add-in manifest issues.</span></span>


## <a name="add-in-dialog-box-cannot-be-displayed"></a><span data-ttu-id="25213-140">La boîte de dialogue des compléments ne s’affiche pas</span><span class="sxs-lookup"><span data-stu-id="25213-140">Add-in dialog box cannot be displayed</span></span>

<span data-ttu-id="25213-p108">Lorsqu’un utilisateur utilise un complément Office, il est invité à autoriser l’affichage d’une boîte de dialogue. L’utilisateur choisit **Autoriser** et le message d’erreur suivant apparaît :</span><span class="sxs-lookup"><span data-stu-id="25213-p108">When using an Office Add-in, the user is asked to allow a dialog box to be displayed. The user chooses **Allow**, and the following error message occurs:</span></span>

<span data-ttu-id="25213-p109">« Les paramètres de sécurité de votre navigateur nous empêchent de créer une boîte de dialogue. Essayez d’utiliser un autre navigateur, ou configurez votre navigateur de sorte que [URL] et le domaine affiché dans la barre d’adresse se trouvent dans la même zone de sécurité. »</span><span class="sxs-lookup"><span data-stu-id="25213-p109">"The security settings in your browser prevent us from creating a dialog box. Try a different browser, or configure your browser so that [URL] and the domain shown in your address bar are in the same security zone."</span></span>

![Capture d’écran du message d’erreur de la boîte de dialogue](http://i.imgur.com/3mqmlgE.png)

|<span data-ttu-id="25213-146">**Navigateurs concernés**</span><span class="sxs-lookup"><span data-stu-id="25213-146">**Affected browsers**</span></span>|<span data-ttu-id="25213-147">**Plateformes concernées**</span><span class="sxs-lookup"><span data-stu-id="25213-147">**Affected platforms**</span></span>|
|:--------------------|:---------------------|
|<span data-ttu-id="25213-148">Internet Explorer, Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="25213-148">Internet Explorer, Microsoft Edge</span></span>|<span data-ttu-id="25213-149">Office sur le web</span><span class="sxs-lookup"><span data-stu-id="25213-149">Office on the web</span></span>|

<span data-ttu-id="25213-p110">Pour résoudre le problème, les utilisateurs finals et les administrateurs peuvent ajouter le domaine du complément à la liste des sites de confiance dans Internet Explorer. Utilisez la même procédure, que vous utilisiez le navigateur Internet Explorer ou Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="25213-p110">To resolve the issue, end users or administrators can add the domain of the add-in to the list of trusted sites in Internet Explorer. Use the same procedure whether you're using the Internet Explorer or Microsoft Edge browser.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="25213-152">n’ajoutez pas l’URL d’un complément à votre liste de sites de confiance si vous ne faites pas confiance au complément.</span><span class="sxs-lookup"><span data-stu-id="25213-152">Do not add the URL for an add-in to your list of trusted sites if you don't trust the add-in.</span></span>

<span data-ttu-id="25213-153">Pour ajouter une URL à votre liste de sites de confiance :</span><span class="sxs-lookup"><span data-stu-id="25213-153">To add a URL to your list of trusted sites:</span></span>

1. <span data-ttu-id="25213-154">Dans Internet Explorer, cliquez sur le bouton Outils et accédez à **Options Internet** > **Sécurité**.</span><span class="sxs-lookup"><span data-stu-id="25213-154">In Internet Explorer, choose the Tools button, and go to **Internet options** > **Security**.</span></span>
2. <span data-ttu-id="25213-155">Sélectionnez la zone **Sites de confiance**, puis choisissez **Sites**.</span><span class="sxs-lookup"><span data-stu-id="25213-155">Select the **Trusted sites** zone, and choose **Sites**.</span></span>
3. <span data-ttu-id="25213-156">Entrez l’URL qui apparaît dans le message d’erreur, puis choisissez **Ajouter**.</span><span class="sxs-lookup"><span data-stu-id="25213-156">Enter the URL that appears in the error message, and choose **Add**.</span></span>
4. <span data-ttu-id="25213-p111">Essayez d’utiliser le complément à nouveau. Si le problème persiste, vérifiez les paramètres pour les autres zones de sécurité et assurez-vous que le domaine du complément se trouve dans la même zone que l’URL qui s’affiche dans la barre d’adresse de l’application Office.</span><span class="sxs-lookup"><span data-stu-id="25213-p111">Try to use the add-in again. If the problem persists, verify the settings for the other security zones and ensure that the add-in domain is in the same zone as the URL that is displayed in the address bar of the Office application.</span></span>

<span data-ttu-id="25213-p112">Ce problème se produit lorsque l’API de la boîte de dialogue est utilisée en mode contextuel. Pour éviter ce problème, utilisez l’indicateur [displayInFrame](/javascript/api/office/office.ui). Cela nécessite que votre page prenne en charge l’affichage dans un iframe. L’exemple suivant montre comment utiliser l’indicateur.</span><span class="sxs-lookup"><span data-stu-id="25213-p112">This issue occurs when the Dialog API is used in pop-up mode. To prevent this issue from occurring, use the [displayInFrame](/javascript/api/office/office.ui) flag. This requires that your page support display within an iframe. The following example shows how to use the flag.</span></span>

```js
Office.context.ui.displayDialogAsync(startAddress, {displayInFrame:true}, callback);
```

## <a name="changes-to-add-in-commands-including-ribbon-buttons-and-menu-items-do-not-take-effect"></a><span data-ttu-id="25213-163">Les modifications apportées aux commandes de complément, y compris les éléments de menu et les boutons du ruban ne s’appliquent pas</span><span class="sxs-lookup"><span data-stu-id="25213-163">Changes to add-in commands including ribbon buttons and menu items do not take effect</span></span>

<span data-ttu-id="25213-164">Si les modifications apportées au manifeste, par exemple aux noms de fichier des icônes de bouton dans le ruban ou au texte des éléments de menu, ne semblent pas appliquées, essayez de vider le cache Office de votre ordinateur.</span><span class="sxs-lookup"><span data-stu-id="25213-164">If changes you've made in the manifest, such as file names of ribbon button icons, or text of add-in commands, do not seem to take effect, try clearing the Office cache on your computer.</span></span> 

#### <a name="for-windows"></a><span data-ttu-id="25213-165">Pour Windows :</span><span class="sxs-lookup"><span data-stu-id="25213-165">For Windows:</span></span>
<span data-ttu-id="25213-166">Supprimer le contenu du dossier `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span><span class="sxs-lookup"><span data-stu-id="25213-166">Delete the content of the folder `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`.</span></span>

#### <a name="for-mac"></a><span data-ttu-id="25213-167">Pour Mac :</span><span class="sxs-lookup"><span data-stu-id="25213-167">For Mac:</span></span>
<span data-ttu-id="25213-168">Supprimer le contenu du dossier `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span><span class="sxs-lookup"><span data-stu-id="25213-168">Delete the content of the folder `~/Library/Containers/com.Microsoft.OsfWebHost/Data/`.</span></span> 

[!include[additional cache folders on Mac](../includes/mac-cache-folders.md)]

#### <a name="for-ios"></a><span data-ttu-id="25213-169">Pour iOS :</span><span class="sxs-lookup"><span data-stu-id="25213-169">For iOS:</span></span>
<span data-ttu-id="25213-p113">Appelez `window.location.reload(true)` à partir de JavaScript dans le complément pour forcer le rechargement. Vous pouvez également choisir de réinstaller Office.</span><span class="sxs-lookup"><span data-stu-id="25213-p113">Call `window.location.reload(true)` from JavaScript in the add-in to force a reload. Alternatively, you can reinstall Office.</span></span>

## <a name="see-also"></a><span data-ttu-id="25213-172">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="25213-172">See also</span></span>

- [<span data-ttu-id="25213-173">Débogage de compléments dans Office sur le web</span><span class="sxs-lookup"><span data-stu-id="25213-173">Debug add-ins in Office on the web</span></span>](debug-add-ins-in-office-online.md) 
- [<span data-ttu-id="25213-174">Charger une version test d’un complément Office sur iPad ou Mac</span><span class="sxs-lookup"><span data-stu-id="25213-174">Sideload an Office Add-in on iPad and Mac</span></span>](sideload-an-office-add-in-on-ipad-and-mac.md)  
- [<span data-ttu-id="25213-175">Débogage des compléments Office sur iPad et Mac</span><span class="sxs-lookup"><span data-stu-id="25213-175">Debug Office Add-ins on iPad and Mac</span></span>](debug-office-add-ins-on-ipad-and-mac.md)  
- [<span data-ttu-id="25213-176">Valider et résoudre des problèmes avec votre manifeste</span><span class="sxs-lookup"><span data-stu-id="25213-176">Validate and troubleshoot issues with your manifest</span></span>](troubleshoot-manifest.md)
    
