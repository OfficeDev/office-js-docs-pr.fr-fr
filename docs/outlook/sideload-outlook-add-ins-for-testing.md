---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 05/13/2021
localization_priority: Normal
ms.openlocfilehash: 9d0fb246f6522c745658a09fce6934ee44d5079a
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555191"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="439aa-103">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="439aa-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="439aa-104">Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="439aa-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-automatically"></a><span data-ttu-id="439aa-105">Sideload automatiquement</span><span class="sxs-lookup"><span data-stu-id="439aa-105">Sideload automatically</span></span>

<span data-ttu-id="439aa-106">Si vous avez créé votre Outlook add-in en utilisant [le générateur Yeoman pour Office Add-ins,](https://github.com/OfficeDev/generator-office)sideloading est préférable de le faire à travers la ligne de commande.</span><span class="sxs-lookup"><span data-stu-id="439aa-106">If you created your Outlook add-in using [the Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office), sideloading is best done through the command line.</span></span> <span data-ttu-id="439aa-107">Cela profitera de notre outillage et de notre charge latérale sur tous vos appareils pris en charge en une seule commande.</span><span class="sxs-lookup"><span data-stu-id="439aa-107">This will take advantage of our tooling and sideload across all of your supported devices in one command.</span></span>

1. <span data-ttu-id="439aa-108">À l’aide de la ligne de commande, accédez à l’annuaire racine de votre projet yeoman généré add-in.</span><span class="sxs-lookup"><span data-stu-id="439aa-108">Using the command line, navigate to the root directory of your Yeoman generated add-in project.</span></span> <span data-ttu-id="439aa-109">Exécutez la commande `npm start`.</span><span class="sxs-lookup"><span data-stu-id="439aa-109">Run the command `npm start`.</span></span>

1. <span data-ttu-id="439aa-110">Votre Outlook add-in sera automatiquement sideload pour Outlook votre ordinateur de bureau.</span><span class="sxs-lookup"><span data-stu-id="439aa-110">Your Outlook add-in will automatically sideload to Outlook on your desktop computer.</span></span> <span data-ttu-id="439aa-111">Vous verrez apparaître un dialogue, indiquant qu’il y a une tentative de sideload l’add-in, énumérant le nom et l’emplacement du fichier manifeste.</span><span class="sxs-lookup"><span data-stu-id="439aa-111">You'll see a dialog appear, stating there is an attempt to sideload the add-in, listing the name and the location of the manifest file.</span></span> <span data-ttu-id="439aa-112">Sélectionnez **OK**, qui enregistrera le manifeste.</span><span class="sxs-lookup"><span data-stu-id="439aa-112">Select **OK**, which will register the manifest.</span></span>

    > [!IMPORTANT]
    > <span data-ttu-id="439aa-113">Si le manifeste contient une erreur ou si le chemin vers le manifeste est invalide, vous recevrez un message d’erreur.</span><span class="sxs-lookup"><span data-stu-id="439aa-113">If the manifest contains an error or the path to the manifest is invalid, you'll receive an error message.</span></span>

1. <span data-ttu-id="439aa-114">Si votre manifeste ne contient aucune erreur et que le chemin est valide, votre module sera désormais sideloaded et disponible à la fois sur votre bureau et dans les Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="439aa-114">If your manifest contains no errors and the path is valid, your add-in will now be sideloaded and available on both your desktop and in Outlook on the web.</span></span> <span data-ttu-id="439aa-115">Il sera également installé sur tous vos appareils pris en charge.</span><span class="sxs-lookup"><span data-stu-id="439aa-115">It will also be installed across all your supported devices.</span></span>

## <a name="sideload-manually"></a><span data-ttu-id="439aa-116">Sideload manuellement</span><span class="sxs-lookup"><span data-stu-id="439aa-116">Sideload manually</span></span>

<span data-ttu-id="439aa-117">Bien que nous vous recommandons fortement de recharger automatiquement à travers la ligne de commande telle que couverte dans la section précédente, vous pouvez également sideload manuellement un add-in Outlook basé sur le client Outlook.</span><span class="sxs-lookup"><span data-stu-id="439aa-117">Though we strongly recommend sideloading automatically through the command line as covered in the previous section, you can also manually sideload an Outlook add-in based on the Outlook client.</span></span>

### <a name="outlook-on-the-web"></a><span data-ttu-id="439aa-118">Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="439aa-118">Outlook on the web</span></span>

<span data-ttu-id="439aa-119">Le processus de chargement latéral d’un add-in Outlook sur le web dépend si vous utilisez la nouvelle version ou classique.</span><span class="sxs-lookup"><span data-stu-id="439aa-119">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="439aa-120">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="439aa-120">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#new-outlook-on-the-web).</span></span>

    ![capture d’écran partielle de la nouvelle version de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="439aa-122">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="439aa-122">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#classic-outlook-on-the-web).</span></span>

    ![capture d’écran partielle de la version classique de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="439aa-124">Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.</span><span class="sxs-lookup"><span data-stu-id="439aa-124">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="new-outlook-on-the-web"></a><span data-ttu-id="439aa-125">Nouvelles Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="439aa-125">New Outlook on the web</span></span>

1. <span data-ttu-id="439aa-126">Accédez à [Outlook sur le web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="439aa-126">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="439aa-127">Créez un nouveau message.</span><span class="sxs-lookup"><span data-stu-id="439aa-127">Create a new message.</span></span>

1. <span data-ttu-id="439aa-128">Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="439aa-128">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Fenêtre de composition de messages dans la nouvelle version d’Outlook sur le web avec l’option pour obtenir des compléments en évidence](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="439aa-130">Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="439aa-130">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Boîte de dialogue Compléments pour Outlook dans la nouvelle version d’Outlook sur le web avec l’option Mes compléments sélectionnée](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="439aa-132">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="439aa-132">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="439aa-133">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="439aa-133">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="439aa-p106">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="439aa-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="classic-outlook-on-the-web"></a><span data-ttu-id="439aa-137">Les Outlook classiques sur le web</span><span class="sxs-lookup"><span data-stu-id="439aa-137">Classic Outlook on the web</span></span>

1. <span data-ttu-id="439aa-138">Accédez à [Outlook sur le web](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="439aa-138">Go to [Outlook on the web](https://outlook.office.com).</span></span>

1. <span data-ttu-id="439aa-139">Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.</span><span class="sxs-lookup"><span data-stu-id="439aa-139">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Capture d’écran d’Outlook sur le web avec une flèche pointant sur l’option Gérer les compléments](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="439aa-141">Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="439aa-141">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Boîte de dialogue du Store Outlook sur le web avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="439aa-143">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="439aa-143">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="439aa-144">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="439aa-144">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="439aa-p108">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="439aa-p108">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-on-the-desktop"></a><span data-ttu-id="439aa-148">Outlook sur le bureau</span><span class="sxs-lookup"><span data-stu-id="439aa-148">Outlook on the desktop</span></span>

#### <a name="outlook-2016-or-later"></a><span data-ttu-id="439aa-149">Outlook 2016 ou plus tard</span><span class="sxs-lookup"><span data-stu-id="439aa-149">Outlook 2016 or later</span></span>

1. <span data-ttu-id="439aa-150">Ouvrez Outlook 2016 ou plus tard sur Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="439aa-150">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="439aa-151">Cliquez sur le bouton **Obtenir des compléments** du ruban.</span><span class="sxs-lookup"><span data-stu-id="439aa-151">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Outlook 2016 ruban pointant vers le bouton Get Add-ins](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="439aa-153">Si vous ne voyez pas le bouton **Get Add-ins** dans votre version de Outlook, sélectionnez :</span><span class="sxs-lookup"><span data-stu-id="439aa-153">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="439aa-154">**Rangez** le bouton sur le ruban, si disponible.</span><span class="sxs-lookup"><span data-stu-id="439aa-154">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="439aa-155">OR</span><span class="sxs-lookup"><span data-stu-id="439aa-155">OR</span></span>
    >
    > - <span data-ttu-id="439aa-156">**Menu** de fichiers, puis sélectionnez **le bouton Manage Add-ins** sur **l’onglet Info** pour ouvrir le dialogue **Add-ins** Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="439aa-156">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="439aa-157">Vous pouvez en savoir plus sur l’expérience Web dans la section [précédente Sideload un add-in dans Outlook sur le web](#outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="439aa-157">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#outlook-on-the-web).</span></span>

1. <span data-ttu-id="439aa-158">S’il y a des onglets près du haut du dialogue, **assurez-vous que l’onglet Add-ins** est sélectionné.</span><span class="sxs-lookup"><span data-stu-id="439aa-158">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="439aa-159">Choisissez **mes add-ins**.</span><span class="sxs-lookup"><span data-stu-id="439aa-159">Choose **My add-ins**.</span></span>

    ![Boîte de dialogue du Store Outlook 2016 avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="439aa-161">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="439aa-161">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="439aa-162">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="439aa-162">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de la page Store avec une flèche pointant vers l’option À partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="439aa-p111">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="439aa-p111">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

#### <a name="outlook-2013"></a><span data-ttu-id="439aa-166">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="439aa-166">Outlook 2013</span></span>

1. <span data-ttu-id="439aa-167">Ouvert Outlook 2013 le Windows.</span><span class="sxs-lookup"><span data-stu-id="439aa-167">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="439aa-168">Sélectionnez **le** menu Fichier, puis **sélectionnez le bouton Manage Add-ins** sur l’onglet **Info.** Outlook ouvrira la version Web dans un navigateur.</span><span class="sxs-lookup"><span data-stu-id="439aa-168">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="439aa-169">Suivez les étapes du [Sideload un add-in dans Outlook sur la](#outlook-on-the-web) section web en fonction de votre version de Outlook sur le web.</span><span class="sxs-lookup"><span data-stu-id="439aa-169">Follow the steps in the [Sideload an add-in in Outlook on the web](#outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="439aa-170">Retirer un add-in sideloaded</span><span class="sxs-lookup"><span data-stu-id="439aa-170">Remove a sideloaded add-in</span></span>

<span data-ttu-id="439aa-171">Sur toutes les versions de Outlook, la clé pour supprimer un module d’ajout sideloaded est le dialogue **My Add-ins** qui répertorie vos modules d’ajout installés. Choisissez l’ellipsis ( `...` ) pour l’add-in puis sélectionnez **Supprimer**.</span><span class="sxs-lookup"><span data-stu-id="439aa-171">On all versions of Outlook, the key to removing a sideloaded add-in is the **My Add-ins** dialog which lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then select **Remove**.</span></span>

<span data-ttu-id="439aa-172">Pour naviguer vers la **boîte de dialogue My Add-ins** pour votre client Outlook, utilisez les dernières étapes répertoriées pour le chargement manuel [dans](#sideload-manually) les sections précédentes de cet article.</span><span class="sxs-lookup"><span data-stu-id="439aa-172">To navigate to the **My Add-ins** dialog box for your Outlook client, use the last steps listed for [manual sideloading](#sideload-manually) in the previous sections of this article.</span></span>

<span data-ttu-id="439aa-173">Pour supprimer un module d’ajout sideloaded de Outlook, utilisez les étapes précédemment décrites dans cet article pour trouver l’add-in dans la section **add-ins personnalisés** de la boîte de dialogue qui répertorie vos modules supplémentaires installés. Choisissez l’ellipsis `...` ( ) pour l’add-in puis **choisissez Supprimer** pour supprimer cet add-in spécifique.</span><span class="sxs-lookup"><span data-stu-id="439aa-173">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the add-in then choose **Remove** to remove that specific add-in.</span></span> <span data-ttu-id="439aa-174">Fermez la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="439aa-174">Close the dialog.</span></span>
