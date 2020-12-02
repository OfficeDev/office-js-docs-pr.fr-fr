---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 12/01/2020
localization_priority: Normal
ms.openlocfilehash: dea2125ccd64eba2e3f1695c8ca1111a710321a4
ms.sourcegitcommit: c2fd7f982f3da748ef6be5c3a7434d859f8b46b9
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/02/2020
ms.locfileid: "49530926"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="533f8-103">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="533f8-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="533f8-104">Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="533f8-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a><span data-ttu-id="533f8-105">Chargement d’un complément dans Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="533f8-105">Sideload an add-in in Outlook on the web</span></span>

<span data-ttu-id="533f8-106">Le processus de chargement d’un complément dans Outlook sur le Web dépend de si vous utilisez la version nouvelle ou classique.</span><span class="sxs-lookup"><span data-stu-id="533f8-106">The process for sideloading an add-in in Outlook on the web depends upon whether you are using the new or classic version.</span></span>

- <span data-ttu-id="533f8-107">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="533f8-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![capture d’écran partielle de la nouvelle version de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="533f8-109">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="533f8-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![capture d’écran partielle de la version classique de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="533f8-111">Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.</span><span class="sxs-lookup"><span data-stu-id="533f8-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="533f8-112">Chargement d’un complément dans la nouvelle version d’Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="533f8-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="533f8-113">Accédez à [Outlook dans Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="533f8-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="533f8-114">Dans Outlook sur le web, créez un message.</span><span class="sxs-lookup"><span data-stu-id="533f8-114">In Outlook on the web, create a new message.</span></span>

1. <span data-ttu-id="533f8-115">Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="533f8-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Fenêtre de composition de messages dans la nouvelle version d’Outlook sur le web avec l’option pour obtenir des compléments en évidence](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="533f8-117">Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="533f8-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Boîte de dialogue Compléments pour Outlook dans la nouvelle version d’Outlook sur le web avec l’option Mes compléments sélectionnée](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="533f8-119">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="533f8-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="533f8-120">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="533f8-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="533f8-p102">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="533f8-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="533f8-124">Chargement d’un complément dans la version classique d’Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="533f8-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="533f8-125">Accédez à [Outlook dans Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="533f8-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="533f8-126">Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.</span><span class="sxs-lookup"><span data-stu-id="533f8-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Capture d’écran d’Outlook sur le web avec une flèche pointant sur l’option Gérer les compléments](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="533f8-128">Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="533f8-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Boîte de dialogue du Store Outlook sur le web avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="533f8-130">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="533f8-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="533f8-131">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="533f8-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="533f8-p104">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="533f8-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="533f8-135">Chargement d’un complément dans la version de bureau d’Outlook</span><span class="sxs-lookup"><span data-stu-id="533f8-135">Sideload an add-in in Outlook on the desktop</span></span>

### <a name="outlook-2016-or-later"></a><span data-ttu-id="533f8-136">Outlook 2016 ou version ultérieure</span><span class="sxs-lookup"><span data-stu-id="533f8-136">Outlook 2016 or later</span></span>

1. <span data-ttu-id="533f8-137">Ouvrez Outlook 2016 ou une version ultérieure sur Windows ou Mac.</span><span class="sxs-lookup"><span data-stu-id="533f8-137">Open Outlook 2016 or later on Windows or Mac.</span></span>

1. <span data-ttu-id="533f8-138">Cliquez sur le bouton **Obtenir des compléments** du ruban.</span><span class="sxs-lookup"><span data-stu-id="533f8-138">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Ruban Outlook 2016 pointant sur le bouton obtenir des compléments](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > <span data-ttu-id="533f8-140">Si vous ne voyez pas le bouton **obtenir des compléments** dans votre version d’Outlook, sélectionnez :</span><span class="sxs-lookup"><span data-stu-id="533f8-140">If you don't see the **Get Add-ins** button in your version of Outlook, select:</span></span>
    >
    > - <span data-ttu-id="533f8-141">Bouton **Store** du ruban, le cas échéant.</span><span class="sxs-lookup"><span data-stu-id="533f8-141">**Store** button on the ribbon, if available.</span></span>
    >
    >   <span data-ttu-id="533f8-142">OR</span><span class="sxs-lookup"><span data-stu-id="533f8-142">OR</span></span>
    >
    > - <span data-ttu-id="533f8-143">Menu **fichier** , puis sélectionnez le bouton **gérer les compléments** sous l’onglet **informations** pour ouvrir la boîte de dialogue **compléments** dans Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="533f8-143">**File** menu, then select the **Manage Add-ins** button on the **Info** tab to open the **Add-ins** dialog in Outlook on the web.</span></span><br><span data-ttu-id="533f8-144">Vous pouvez en savoir plus sur l’expérience Web dans la section précédente [chargement d’un complément dans Outlook sur le Web](#sideload-an-add-in-in-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="533f8-144">You can see more about the web experience in the previous section [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web).</span></span>

1. <span data-ttu-id="533f8-145">S’il y a des onglets dans la partie supérieure de la boîte de dialogue, vérifiez que l’onglet **compléments** est sélectionné.</span><span class="sxs-lookup"><span data-stu-id="533f8-145">If there are tabs near the top of the dialog, ensure that the **Add-ins** tab is selected.</span></span> <span data-ttu-id="533f8-146">Sélectionnez **mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="533f8-146">Choose **My add-ins**.</span></span>

    ![Boîte de dialogue du Store Outlook 2016 avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="533f8-148">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="533f8-148">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="533f8-149">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="533f8-149">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de la page Store avec une flèche pointant vers l’option À partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="533f8-p107">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="533f8-p107">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="outlook-2013"></a><span data-ttu-id="533f8-153">Outlook 2013</span><span class="sxs-lookup"><span data-stu-id="533f8-153">Outlook 2013</span></span>

1. <span data-ttu-id="533f8-154">Ouvrez Outlook 2013 sur Windows.</span><span class="sxs-lookup"><span data-stu-id="533f8-154">Open Outlook 2013 on Windows.</span></span>

1. <span data-ttu-id="533f8-155">Sélectionnez le menu **fichier** , puis cliquez sur le bouton **gérer les compléments** sous l’onglet **informations** . Outlook ouvre la version Web dans un navigateur.</span><span class="sxs-lookup"><span data-stu-id="533f8-155">Select the **File** menu, then select the **Manage Add-ins** button on the **Info** tab. Outlook will open the web version in a browser.</span></span>

1. <span data-ttu-id="533f8-156">Suivez les étapes de la section [chargement d’un complément dans Outlook sur le Web](#sideload-an-add-in-in-outlook-on-the-web) en fonction de votre version d’Outlook sur le Web.</span><span class="sxs-lookup"><span data-stu-id="533f8-156">Follow the steps in the [Sideload an add-in in Outlook on the web](#sideload-an-add-in-in-outlook-on-the-web) section according to your version of Outlook on the web.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="533f8-157">Supprimer un complément versions test chargées</span><span class="sxs-lookup"><span data-stu-id="533f8-157">Remove a sideloaded add-in</span></span>

<span data-ttu-id="533f8-158">Pour supprimer un complément versions test chargées à partir d’Outlook, suivez les étapes décrites précédemment dans cet article pour trouver le complément dans la section **compléments personnalisés** de la boîte de dialogue qui répertorie vos compléments installés. Sélectionnez les points de suspension ( `...` ) pour le complément, puis cliquez sur **supprimer** pour supprimer ce complément spécifique.</span><span class="sxs-lookup"><span data-stu-id="533f8-158">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>