---
title: Déboguer votre Outlook d’événement (prévisualisation)
description: Découvrez comment déboguer votre complément Outlook qui implémente l’activation basée sur des événements.
ms.topic: article
ms.date: 05/14/2021
localization_priority: Normal
ms.openlocfilehash: d7621a7407db3b8e773d1534beb6c881f7b48558
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555268"
---
# <a name="debug-your-event-based-outlook-add-in-preview"></a><span data-ttu-id="8a581-103">Déboguer votre Outlook d’événement (prévisualisation)</span><span class="sxs-lookup"><span data-stu-id="8a581-103">Debug your event-based Outlook add-in (preview)</span></span>

<span data-ttu-id="8a581-104">Cet article fournit des instructions de débogage lorsque vous implémentez l’activation basée sur des [événements](autolaunch.md) dans votre complément.</span><span class="sxs-lookup"><span data-stu-id="8a581-104">This article provides debugging guidance as you implement [event-based activation](autolaunch.md) in your add-in.</span></span> <span data-ttu-id="8a581-105">La fonctionnalité d’activation basée sur des événements est actuellement en prévisualisation.</span><span class="sxs-lookup"><span data-stu-id="8a581-105">The event-based activation feature is currently in preview.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="8a581-106">Cette fonctionnalité de débogage est uniquement prise en charge pour la prévisualisation dans Outlook sur Windows avec un abonnement Microsoft 365'abonnement.</span><span class="sxs-lookup"><span data-stu-id="8a581-106">This debugging capability is only supported for preview in Outlook on Windows with a Microsoft 365 subscription.</span></span> <span data-ttu-id="8a581-107">Pour plus d’informations, voir la section [Débogage d’aperçu](#preview-debugging-for-the-event-based-activation-feature) pour la fonctionnalité d’activation basée sur des événements dans cet article.</span><span class="sxs-lookup"><span data-stu-id="8a581-107">For more information, see the [Preview debugging for the event-based activation feature](#preview-debugging-for-the-event-based-activation-feature) section in this article.</span></span>

<span data-ttu-id="8a581-108">Dans cet article, nous abordons les étapes clés pour activer le débogage.</span><span class="sxs-lookup"><span data-stu-id="8a581-108">In this article, we discuss the key stages to enable debugging.</span></span>

- [<span data-ttu-id="8a581-109">Marquer le add-in pour le débogage</span><span class="sxs-lookup"><span data-stu-id="8a581-109">Mark the add-in for debugging</span></span>](#mark-your-add-in-for-debugging)
- [<span data-ttu-id="8a581-110">Configurer Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="8a581-110">Configure Visual Studio Code</span></span>](#configure-visual-studio-code)
- [<span data-ttu-id="8a581-111">Attacher les Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="8a581-111">Attach Visual Studio Code</span></span>](#attach-visual-studio-code)
- [<span data-ttu-id="8a581-112">Debug</span><span class="sxs-lookup"><span data-stu-id="8a581-112">Debug</span></span>](#debug)

<span data-ttu-id="8a581-113">Plusieurs options s’offrent à vous pour créer votre projet de add-in.</span><span class="sxs-lookup"><span data-stu-id="8a581-113">You have several options for creating your add-in project.</span></span> <span data-ttu-id="8a581-114">En fonction de l’option que vous utilisez, les étapes peuvent varier.</span><span class="sxs-lookup"><span data-stu-id="8a581-114">Depending on the option you're using, the steps may vary.</span></span> <span data-ttu-id="8a581-115">Si c’est le cas, si vous avez utilisé le générateur Yeoman pour les compléments Office pour créer votre projet de complément (par exemple, en  faisant la procédure pas à pas [d’activation](autolaunch.md)basée sur l’événement), suivez les étapes de **yo office,** sinon suivez les autres étapes.</span><span class="sxs-lookup"><span data-stu-id="8a581-115">Where this is the case, if you used the Yeoman generator for Office Add-ins to create your add-in project (for example, by doing the [event-based activation walkthrough](autolaunch.md)), then follow the **yo office** steps, otherwise follow the **Other** steps.</span></span> <span data-ttu-id="8a581-116">Visual Studio Code doit être au moins la version 1.56.1.</span><span class="sxs-lookup"><span data-stu-id="8a581-116">Visual Studio Code should be at least version 1.56.1.</span></span>

## <a name="preview-debugging-for-the-event-based-activation-feature"></a><span data-ttu-id="8a581-117">Prévisualiser le débogage pour la fonctionnalité d’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="8a581-117">Preview debugging for the event-based activation feature</span></span>

<span data-ttu-id="8a581-118">Nous vous invitons à tester la fonctionnalité de débogage pour la fonctionnalité d’activation basée sur des événements !</span><span class="sxs-lookup"><span data-stu-id="8a581-118">We invite you to try out the debugging capability for the event-based activation feature!</span></span> <span data-ttu-id="8a581-119">Faites-nous part de vos scénarios et de la façon dont nous pouvons les améliorer en nous faisant part de vos commentaires GitHub (voir la **section** Commentaires à la fin de cette page).</span><span class="sxs-lookup"><span data-stu-id="8a581-119">Let us know your scenarios and how we can improve by giving us feedback through GitHub (see the **Feedback** section at the end of this page).</span></span>

<span data-ttu-id="8a581-120">Pour prévisualiser cette fonctionnalité Outlook sur Windows, la version minimale requise est 16.0.13729.20000.</span><span class="sxs-lookup"><span data-stu-id="8a581-120">To preview this capability for Outlook on Windows, the minimum required build is 16.0.13729.20000.</span></span> <span data-ttu-id="8a581-121">Pour accéder à Office versions bêta, rejoignez [le programme Office Insider.](https://insider.office.com)</span><span class="sxs-lookup"><span data-stu-id="8a581-121">For access to Office beta builds, join the [Office Insider program](https://insider.office.com).</span></span>

## <a name="mark-your-add-in-for-debugging"></a><span data-ttu-id="8a581-122">Marquer votre add-in pour le débogage</span><span class="sxs-lookup"><span data-stu-id="8a581-122">Mark your add-in for debugging</span></span>

1. <span data-ttu-id="8a581-123">Définissez la clé de `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` Registre.</span><span class="sxs-lookup"><span data-stu-id="8a581-123">Set the registry key `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span> <span data-ttu-id="8a581-124">`[Add-in ID]` est **l’ID** dans le manifeste du add-in.</span><span class="sxs-lookup"><span data-stu-id="8a581-124">`[Add-in ID]` is the **Id** in the add-in manifest.</span></span>

    <span data-ttu-id="8a581-125">**yo office**: dans une fenêtre de ligne de commande, accédez à la racine du dossier de votre add-in, puis exécutez la commande suivante.</span><span class="sxs-lookup"><span data-stu-id="8a581-125">**yo office**: In a command line window, navigate to the root of your add-in folder then run the following command.</span></span>

    ```command&nbsp;line
    npm start
    ```

    <span data-ttu-id="8a581-126">Outre la création du code et le démarrage du serveur local, cette commande doit définir la clé de Registre pour `UseDirectDebugger` ce complément sur `1` .</span><span class="sxs-lookup"><span data-stu-id="8a581-126">In addition to building the code and starting the local server, this command should set the `UseDirectDebugger` registry key for this add-in to `1`.</span></span>

    <span data-ttu-id="8a581-127">**Autre**: ajoutez la `UseDirectDebugger` clé de Registre sous `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\` .</span><span class="sxs-lookup"><span data-stu-id="8a581-127">**Other**: Add the `UseDirectDebugger` registry key under `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\Developer\[Add-in ID]\`.</span></span> <span data-ttu-id="8a581-128">Remplacez `[Add-in ID]` par **l’ID** du manifeste du module.</span><span class="sxs-lookup"><span data-stu-id="8a581-128">Replace `[Add-in ID]` with the **Id** from the add-in manifest.</span></span> <span data-ttu-id="8a581-129">Définissez la clé de Registre sur `1` .</span><span class="sxs-lookup"><span data-stu-id="8a581-129">Set the registry key to `1`.</span></span>

    [!include[Developer registry key](../includes/developer-registry-key.md)]

1. <span data-ttu-id="8a581-130">Démarrez Outlook bureau (ou redémarrez Outlook s’il est déjà ouvert).</span><span class="sxs-lookup"><span data-stu-id="8a581-130">Start Outlook desktop (or restart Outlook if it's already open).</span></span>
1. <span data-ttu-id="8a581-131">Rédigez un nouveau message ou rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="8a581-131">Compose a new message or appointment.</span></span> <span data-ttu-id="8a581-132">Vous devriez voir la boîte de dialogue suivante.</span><span class="sxs-lookup"><span data-stu-id="8a581-132">You should see the following dialog.</span></span> <span data-ttu-id="8a581-133">*N’interagissez* pas encore avec la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="8a581-133">Do *not* interact with the dialog yet.</span></span>

    ![Capture d’écran de la boîte de dialogue Debug Event-based handler](../images/outlook-win-autolaunch-debug-dialog.png)

## <a name="configure-visual-studio-code"></a><span data-ttu-id="8a581-135">Configurer Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="8a581-135">Configure Visual Studio Code</span></span>

### <a name="yo-office"></a><span data-ttu-id="8a581-136">yo office</span><span class="sxs-lookup"><span data-stu-id="8a581-136">yo office</span></span>

1. <span data-ttu-id="8a581-137">De retour dans la fenêtre de ligne de commande, ouvrez Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="8a581-137">Back in the command line window, open Visual Studio Code.</span></span>

    ```command&nbsp;line
    code .
    ```

1. <span data-ttu-id="8a581-138">Dans Visual Studio Code, ouvrez le fichier **./.vscode/launch.js** et ajoutez l’extrait suivant à votre liste de configurations.</span><span class="sxs-lookup"><span data-stu-id="8a581-138">In Visual Studio Code, open the file **./.vscode/launch.json** and add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="8a581-139">Enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="8a581-139">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

### <a name="other"></a><span data-ttu-id="8a581-140">Autre</span><span class="sxs-lookup"><span data-stu-id="8a581-140">Other</span></span>

1. <span data-ttu-id="8a581-141">Créez un dossier appelé **Débogage** (éventuellement dans votre **dossier Bureau).**</span><span class="sxs-lookup"><span data-stu-id="8a581-141">Create a new folder called **Debugging** (perhaps in your **Desktop** folder).</span></span>
1. <span data-ttu-id="8a581-142">Ouvrez Visual Studio Code.</span><span class="sxs-lookup"><span data-stu-id="8a581-142">Open Visual Studio Code.</span></span>
1. <span data-ttu-id="8a581-143">Accédez **à Dossier**  >  **d’ouverture de** fichier, accédez au dossier que vous avez créé, puis **sélectionnez Sélectionner un dossier.**</span><span class="sxs-lookup"><span data-stu-id="8a581-143">Go to **File** > **Open Folder**, navigate to the folder you just created, then choose **Select Folder**.</span></span>
1. <span data-ttu-id="8a581-144">Dans la barre d’activité, sélectionnez **l’élément Débogage** (Ctrl+Shift+D).</span><span class="sxs-lookup"><span data-stu-id="8a581-144">On the Activity Bar, select the **Debug** item (Ctrl+Shift+D).</span></span>

    ![Capture d’écran de l’icône Débogage dans la barre d’activité](../images/vs-code-debug.png)

1. <span data-ttu-id="8a581-146">Sélectionnez **créer une launch.jssur le lien de** fichier.</span><span class="sxs-lookup"><span data-stu-id="8a581-146">Select the **create a launch.json file** link.</span></span>

    ![Capture d’écran du lien pour créer une launch.jsfichier dans Visual Studio Code](../images/vs-code-create-launch.json.png)

1. <span data-ttu-id="8a581-148">Dans la **dropdown Sélectionner un** environnement, **sélectionnez Edge : Lancer** pour créer une launch.jsfichier.</span><span class="sxs-lookup"><span data-stu-id="8a581-148">In the **Select Environment** dropdown, select **Edge: Launch** to create a launch.json file.</span></span>
1. <span data-ttu-id="8a581-149">Ajoutez l’extrait suivant à votre liste de configurations.</span><span class="sxs-lookup"><span data-stu-id="8a581-149">Add the following excerpt to your list of configurations.</span></span> <span data-ttu-id="8a581-150">Enregistrez vos modifications.</span><span class="sxs-lookup"><span data-stu-id="8a581-150">Save your changes.</span></span>

    ```json
    {
      "name": "Direct Debugging",
      "type": "node",
      "request": "attach",
      "port": 9229,
      "protocol": "inspector",
      "timeout": 600000,
      "trace": true
    }
    ```

## <a name="attach-visual-studio-code"></a><span data-ttu-id="8a581-151">Attacher les Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="8a581-151">Attach Visual Studio Code</span></span>

1. <span data-ttu-id="8a581-152">Pour rechercher l’ID **dubundle.js,** ouvrez le dossier suivant dans l’Explorateur Windows et recherchez l’ID de votre Windows (trouvé dans le manifeste). </span><span class="sxs-lookup"><span data-stu-id="8a581-152">To find the add-in's **bundle.js**, open the following folder in Windows Explorer and search for your add-in's **Id** (found in the manifest).</span></span>

    ```text
    %LOCALAPPDATA%\Microsoft\Office\16.0\Wef
    ```

    <span data-ttu-id="8a581-153">Ouvrez le dossier précédé de cet ID et copiez son chemin d’accès complet.</span><span class="sxs-lookup"><span data-stu-id="8a581-153">Open the folder prefixed with this ID and copy its full path.</span></span> <span data-ttu-id="8a581-154">Dans Visual Studio Code, ouvrez **bundle.js** à partir de ce dossier.</span><span class="sxs-lookup"><span data-stu-id="8a581-154">In Visual Studio Code, open **bundle.js** from that folder.</span></span> <span data-ttu-id="8a581-155">Le modèle du chemin d’accès au fichier doit être le suivant :</span><span class="sxs-lookup"><span data-stu-id="8a581-155">The pattern of the file path should be as follows:</span></span>

    `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\{[Outlook profile GUID]}\[encoding]\Javascript\[Add-in ID]_[Add-in Version]_[locale]\bundle.js`

1. <span data-ttu-id="8a581-156">Placez les points d’arrêt bundle.js l’endroit où vous souhaitez que le débogger s’arrête.</span><span class="sxs-lookup"><span data-stu-id="8a581-156">Place breakpoints in bundle.js where you want the debugger to stop.</span></span>
1. <span data-ttu-id="8a581-157">Dans la **dropdown DEBUG,** sélectionnez le nom **Débogage** direct, puis sélectionnez **Exécuter**.</span><span class="sxs-lookup"><span data-stu-id="8a581-157">In the **DEBUG** dropdown, select the name **Direct Debugging**, then select **Run**.</span></span>

    ![Capture d’écran de la sélection du débogage direct à partir des options de configuration dans la Visual Studio Code de débogage](../images/outlook-win-autolaunch-debug-vsc.png)

## <a name="debug"></a><span data-ttu-id="8a581-159">Debug</span><span class="sxs-lookup"><span data-stu-id="8a581-159">Debug</span></span>

1. <span data-ttu-id="8a581-160">Après avoir confirmé que le déboguer est attaché, revenir  à Outlook, puis dans la boîte de dialogue de débogage basée sur l’événement, choisissez **OK** .</span><span class="sxs-lookup"><span data-stu-id="8a581-160">After confirming that the debugger is attached, return to Outlook, and in the **Debug Event-based handler** dialog, choose **OK** .</span></span>

1. <span data-ttu-id="8a581-161">Vous pouvez désormais atteindre vos points d’arrêt dans Visual Studio Code, ce qui vous permet de déboguer votre code d’activation basé sur des événements.</span><span class="sxs-lookup"><span data-stu-id="8a581-161">You can now hit your breakpoints in Visual Studio Code, enabling you to debug your event-based activation code.</span></span>

## <a name="stop-debugging"></a><span data-ttu-id="8a581-162">Arrêter le débogage</span><span class="sxs-lookup"><span data-stu-id="8a581-162">Stop debugging</span></span>

<span data-ttu-id="8a581-163">Pour arrêter le débogage pour le reste de la session de bureau Outlook en cours, dans la boîte de dialogue **Debug Event-based handler** ( Annuler ).</span><span class="sxs-lookup"><span data-stu-id="8a581-163">To stop debugging for the rest of the current Outlook desktop session, in the **Debug Event-based handler** dialog, choose **Cancel**.</span></span> <span data-ttu-id="8a581-164">Pour ré-activer le débogage, redémarrez Outlook bureau.</span><span class="sxs-lookup"><span data-stu-id="8a581-164">To re-enable debugging, restart Outlook desktop.</span></span>

<span data-ttu-id="8a581-165">Pour empêcher que la boîte de dialogue du **handler** basé sur un événement de débogage s’insérable et arrêter le débogage pour les sessions Outlook suivantes, supprimez la clé de Registre associée ou définissez sa valeur sur : `0` `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger` .</span><span class="sxs-lookup"><span data-stu-id="8a581-165">To prevent the **Debug Event-based handler** dialog from popping up and stop debugging for subsequent Outlook sessions, delete the associated registry key or set its value to `0`: `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Wef\Developer\[Add-in ID]\UseDirectDebugger`.</span></span>

## <a name="see-also"></a><span data-ttu-id="8a581-166">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="8a581-166">See also</span></span>

- [<span data-ttu-id="8a581-167">Configurer votre complément Outlook pour l’activation basée sur des événements</span><span class="sxs-lookup"><span data-stu-id="8a581-167">Configure your Outlook add-in for event-based activation</span></span>](autolaunch.md)
- [<span data-ttu-id="8a581-168">Déboguer votre complément avec la journalisation runtime</span><span class="sxs-lookup"><span data-stu-id="8a581-168">Debug your add-in with runtime logging</span></span>](../testing/runtime-logging.md#runtime-logging-on-windows)
