---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 47eb5da19f858b6e30339acc59da24a818fc0959
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/23/2021
ms.locfileid: "53077028"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="64b64-103">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="64b64-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="64b64-104">Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="64b64-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="64b64-105">Chargement de version de version de version automatique</span><span class="sxs-lookup"><span data-stu-id="64b64-105">Sideload automatically</span></span>

<span data-ttu-id="64b64-106">Si vous avez créé votre Outlook à l’aide du générateur [Yeoman](https://github.com/OfficeDev/generator-office)pour les Office, il est préférable de faire un chargement de version de version par le biais de la ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="64b64-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="64b64-107">Cela tirera parti de nos outils et de notre chargement de version de version sur tous vos appareils pris en charge dans une seule commande.</span><span class="sxs-lookup"><span data-stu-id="64b64-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="64b64-108">À l’aide de la ligne de commande, accédez au répertoire racine de votre projet de add-in généré par Yeoman.</span><span class="sxs-lookup"><span data-stu-id="64b64-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="64b64-109">Exécutez la commande `npm start`.</span><span class="sxs-lookup"><span data-stu-id="64b64-109">Run the command `npm start`.</span></span>

1. <span data-ttu-id="64b64-110">Votre Outlook de bureau est automatiquement chargé de manière Outlook sur votre ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="64b64-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="64b64-111">Une boîte de dialogue s’affiche, indiquant qu’il y a une tentative de chargement de version de chargement du module, répertoriant le nom et l’emplacement du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="64b64-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="64b64-112">Sélectionnez **OK,** qui enregistre le manifeste.</span><span class="sxs-lookup"><span data-stu-id="64b64-112">Select **OK**, which will register the manifest.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="64b64-113">Si le manifeste contient une erreur ou si le chemin d’accès au manifeste n’est pas valide, vous recevrez un message d’erreur.</span><span class="sxs-lookup"><span data-stu-id="64b64-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

1. <span data-ttu-id="64b64-114">Si votre manifeste ne contient pas d’erreurs et que le chemin d’accès est valide, votre application est désormais rechargée de côté et disponible à la fois sur votre bureau et dans Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="64b64-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="64b64-115">Il sera également installé sur tous vos appareils pris en charge.</span><span class="sxs-lookup"><span data-stu-id="64b64-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="64b64-116">Chargement de version de version manuelle</span><span class="sxs-lookup"><span data-stu-id="64b64-116">Sideload manually</span></span>

<span data-ttu-id="64b64-117">Bien que nous recommandions vivement le chargement d’une version de version secondaire automatiquement par le biais de la ligne de commande comme abordé dans la section précédente, vous pouvez également charger manuellement une version de version de chargement de version de version antérieure d’un Outlook basé sur le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="64b64-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="64b64-118">Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="64b64-118">Outlook on the web</span></span>

<span data-ttu-id="64b64-119">Le processus de chargement d’une version de version Outlook sur le web dépend de l’utilisation de la nouvelle version ou de la version classique.</span><span class="sxs-lookup"><span data-stu-id="64b64-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="64b64-120">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="64b64-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![Capture d’écran partielle de la nouvelle barre Outlook sur le web’outils.](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="64b64-122">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="64b64-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![Capture d’écran partielle de la barre d’outils Outlook sur le web classique.](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="64b64-124">Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.</span><span class="sxs-lookup"><span data-stu-id="64b64-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="64b64-125">Nouvelle Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="64b64-125">New Outlook on the web</span></span>

1. <span data-ttu-id="64b64-126">Accédez à [Outlook sur le web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="64b64-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="64b64-127">Créez un message.</span><span class="sxs-lookup"><span data-stu-id="64b64-127">Create a new message.</span></span>

1. <span data-ttu-id="64b64-128">Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="64b64-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Fenêtre de composition de message dans la nouvelle Outlook sur le web avec l’option Obtenir des add-ins mise en évidence.](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="64b64-130">Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="64b64-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Les applications pour Outlook boîte de dialogue dans la nouvelle Outlook sur le web avec Mes applications sélectionnées.](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="64b64-132">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="64b64-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="64b64-133">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="64b64-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran gérer les add-ins pointant vers Ajouter à partir d’une option de fichier.](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="64b64-p106">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="64b64-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="64b64-137">Modèle Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="64b64-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="64b64-138">Accédez à [Outlook sur le web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="64b64-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="64b64-139">Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.</span><span class="sxs-lookup"><span data-stu-id="64b64-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Outlook sur le web capture d’écran pointant vers l’option Gérer les add-ins.](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="64b64-141">Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="64b64-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Outlook sur le web dans la boîte de dialogue Du store avec mes applications sélectionnées.](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="64b64-143">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="64b64-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="64b64-144">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="64b64-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran gérer les add-ins pointant vers Ajouter à partir d’une option de fichier.](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="64b64-p108">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="64b64-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="64b64-148">Outlook sur le bureau</span><span class="sxs-lookup"><span data-stu-id="64b64-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="64b64-149">Outlook 2016 ou ultérieure</span><span class="sxs-lookup"><span data-stu-id="64b64-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="64b64-150">Ouvrez Outlook 2016 ou ultérieurement sur Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="64b64-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="64b64-151">Cliquez sur le bouton **Obtenir des compléments** du ruban.</span><span class="sxs-lookup"><span data-stu-id="64b64-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Outlook 2016 ruban pointant vers le bouton Obtenir des modules.](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="64b64-153">Si vous ne voyez pas le bouton Obtenir **des** Outlook, sélectionnez :</span><span class="sxs-lookup"><span data-stu-id="64b64-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="64b64-154">**Bouton Stocker** sur le ruban, si disponible.</span><span class="sxs-lookup"><span data-stu-id="64b64-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="64b64-155">OU</span><span class="sxs-lookup"><span data-stu-id="64b64-155">OR</span></span>
    >
    > - <span data-ttu-id="64b64-156">**Menu** Fichier, puis sélectionnez le bouton Gérer les **modules complémentaires** sous l’onglet **Informations** pour ouvrir la boîte de dialogue Des Outlook sur le web. </span><span class="sxs-lookup"><span data-stu-id="64b64-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="64b64-157">Vous pouvez en savoir plus sur l’expérience web dans la section précédente chargement de version de chargement d’un [Outlook sur le web](#outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="64b64-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="64b64-158">S’il existe des onglets en haut de la boîte de dialogue, **assurez-vous** que l’onglet Des applications est sélectionné.</span><span class="sxs-lookup"><span data-stu-id="64b64-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="64b64-159">Choose **My add-ins**.</span><span class="sxs-lookup"><span data-stu-id="64b64-159">Choose **My add-ins**.</span></span>

    ![Outlook 2016 dans la boîte de dialogue Du store avec mes applications sélectionnées.](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="64b64-161">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="64b64-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="64b64-162">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="64b64-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran du magasin pointant sur Ajouter à partir d’une option de fichier.](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="64b64-p111">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="64b64-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="64b64-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="64b64-166">Outlook 2013</span></span>

1. <span data-ttu-id="64b64-167">Ouvrez Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="64b64-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="64b64-168">Sélectionnez **le** menu Fichier, puis sélectionnez le bouton Gérer les **modules complémentaires** sous l’onglet **Informations.** Outlook ouvre la version web dans un navigateur.</span><span class="sxs-lookup"><span data-stu-id="64b64-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="64b64-169">Suivez les étapes de la section Chargement de version de [version](#outlook-on-the-web) Outlook sur le web en fonction de votre version de Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="64b64-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="64b64-170">Supprimer un add-in chargé de nouveau</span><span class="sxs-lookup"><span data-stu-id="64b64-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="64b64-171">Sur toutes les versions de Outlook, la clé de la suppression  d’un module de chargement de version ultérieure est la boîte de dialogue Mes applications qui répertorie vos applications installées. Choisissez les ellipses ( ) pour le `...` add-in, puis sélectionnez **Supprimer**.</span><span class="sxs-lookup"><span data-stu-id="64b64-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="64b64-172">Pour accéder à la boîte de dialogue Mes applications pour votre client Outlook, [](#sideload-manually) utilisez les dernières **étapes** répertoriées pour le chargement de version manuelle dans les sections précédentes de cet article.</span><span class="sxs-lookup"><span data-stu-id="64b64-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="64b64-173">Pour supprimer un **add-in** chargé de Outlook, utilisez les étapes décrites précédemment dans cet article pour rechercher le module dans la section Des applications personnalisées de la boîte de dialogue répertoriant vos applications installées. Choisissez les ellipses ( ) pour le module, puis choisissez Supprimer pour `...` supprimer ce dernier. </span><span class="sxs-lookup"><span data-stu-id="64b64-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="64b64-174">Fermez la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="64b64-174">Close the dialog.</span></span>
