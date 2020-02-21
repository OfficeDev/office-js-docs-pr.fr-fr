---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 06/24/2019
localization_priority: Normal
ms.openlocfilehash: b177e6adbc4ac702b7bd9dcec38f2fe2d2f29cf1
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166076"
---
# <a name="sideload-outlook-add-ins-for-testing"></a><span data-ttu-id="8ee69-103">Chargement de version test des compléments Outlook</span><span class="sxs-lookup"><span data-stu-id="8ee69-103">Sideload Outlook add-ins for testing</span></span>

<span data-ttu-id="8ee69-104">Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.</span><span class="sxs-lookup"><span data-stu-id="8ee69-104">You can use sideloading to install an Outlook add-in for testing without having to first put it in an add-in catalog.</span></span>


## <a name="sideload-an-add-in-in-outlook-in-office-365"></a><span data-ttu-id="8ee69-105">Chargement d’une version test d’un complément dans Outlook dans Office 365</span><span class="sxs-lookup"><span data-stu-id="8ee69-105">Sideload an add-in in Outlook in Office 365</span></span>

<span data-ttu-id="8ee69-106">Le processus de chargement de la version test d’un complément dans Outlook dans Office 365 dépend de si vous utilisez la nouvelle version d’Outlook sur le web ou la version classique.</span><span class="sxs-lookup"><span data-stu-id="8ee69-106">The process for sideloading an add-in in Outlook in Office 365 depends upon whether you are using the new Outlook on the web or classic Outlook on the web.</span></span>

- <span data-ttu-id="8ee69-107">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="8ee69-107">If your mailbox toolbar looks like the following image, see [Sideload an add-in in the new Outlook on the web](#sideload-an-add-in-in-the-new-outlook-on-the-web).</span></span>

    ![capture d’écran partielle de la nouvelle version de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-new-toolbar.png)

- <span data-ttu-id="8ee69-109">Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span><span class="sxs-lookup"><span data-stu-id="8ee69-109">If your mailbox toolbar looks like the following image, see [Sideload an add-in in classic Outlook on the web](#sideload-an-add-in-in-classic-outlook-on-the-web).</span></span>

    ![capture d’écran partielle de la version classique de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> <span data-ttu-id="8ee69-111">Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.</span><span class="sxs-lookup"><span data-stu-id="8ee69-111">If your organization has included its logo in the mailbox toolbar, you might see something slightly different than shown in the preceding images.</span></span>

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a><span data-ttu-id="8ee69-112">Chargement d’un complément dans la nouvelle version d’Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="8ee69-112">Sideload an add-in in the new Outlook on the web</span></span>

1. <span data-ttu-id="8ee69-113">Accédez à [Outlook dans Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="8ee69-113">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="8ee69-114">Dans Outlook sur le web, créez un message.</span><span class="sxs-lookup"><span data-stu-id="8ee69-114">In Outlook on the web, create a new message.</span></span>   

1. <span data-ttu-id="8ee69-115">Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.</span><span class="sxs-lookup"><span data-stu-id="8ee69-115">Choose **...** from the bottom of the new message and then select **Get Add-ins** from the menu that appears.</span></span>

    ![Fenêtre de composition de messages dans la nouvelle version d’Outlook sur le web avec l’option pour obtenir des compléments en évidence](../images/outlook-on-the-web-new-get-add-ins.png)

1. <span data-ttu-id="8ee69-117">Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="8ee69-117">In the **Add-Ins for Outlook** dialog box, select **My add-ins**.</span></span>

    ![Boîte de dialogue Compléments pour Outlook dans la nouvelle version d’Outlook sur le web avec l’option Mes compléments sélectionnée](../images/outlook-on-the-web-new-my-add-ins.png)

1. <span data-ttu-id="8ee69-119">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="8ee69-119">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="8ee69-120">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="8ee69-120">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="8ee69-p102">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="8ee69-p102">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a><span data-ttu-id="8ee69-124">Chargement d’un complément dans la version classique d’Outlook sur le web</span><span class="sxs-lookup"><span data-stu-id="8ee69-124">Sideload an add-in in classic Outlook on the web</span></span>

1. <span data-ttu-id="8ee69-125">Accédez à [Outlook dans Office 365](https://outlook.office.com).</span><span class="sxs-lookup"><span data-stu-id="8ee69-125">Go to [Outlook in Office 365](https://outlook.office.com).</span></span>

1. <span data-ttu-id="8ee69-126">Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.</span><span class="sxs-lookup"><span data-stu-id="8ee69-126">Choose the gear icon in the top-right section of the toolbar and select **Manage add-ins**.</span></span>

    ![Capture d’écran d’Outlook sur le web avec une flèche pointant sur l’option Gérer les compléments](../images/outlook-sideload-web-manage-integrations.png)

1. <span data-ttu-id="8ee69-128">Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="8ee69-128">On the **Manage add-ins** page, select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Boîte de dialogue du Store Outlook sur le web avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="8ee69-130">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="8ee69-130">Locate the **Custom add-ins** section at the bottom of the dialog box.</span></span> <span data-ttu-id="8ee69-131">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="8ee69-131">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="8ee69-p104">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="8ee69-p104">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a><span data-ttu-id="8ee69-135">Chargement d’un complément dans la version de bureau d’Outlook</span><span class="sxs-lookup"><span data-stu-id="8ee69-135">Sideload an add-in in Outlook on the desktop</span></span>

1. <span data-ttu-id="8ee69-136">Ouvrez Outlook 2013 ou une version ultérieure sur Windows, ou Outlook 2016 ou une version ultérieure sur Mac.</span><span class="sxs-lookup"><span data-stu-id="8ee69-136">Open Outlook 2013 or later on Windows, or Outlook 2016 or later on Mac.</span></span>

1. <span data-ttu-id="8ee69-137">Cliquez sur le bouton **Obtenir des compléments** du ruban.</span><span class="sxs-lookup"><span data-stu-id="8ee69-137">Select the **Get Add-ins** button on the ribbon.</span></span>

    ![Ruban Outlook 2016 avec une flèche pointant sur le bouton Store](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > <span data-ttu-id="8ee69-139">Si vous ne voyez pas le bouton **Obtenir des compléments** dans votre version d’Outlook, cliquez sur le bouton **Store** situé dans le ruban à la place.</span><span class="sxs-lookup"><span data-stu-id="8ee69-139">If you don't see the **Get Add-ins** button in your version of Outlook, select the **Store** button on the ribbon instead.</span></span>

1. <span data-ttu-id="8ee69-140">Sélectionnez **Compléments**, puis **Mes compléments**.</span><span class="sxs-lookup"><span data-stu-id="8ee69-140">Select **Add-Ins**, and then select **My add-ins**.</span></span>

    ![Boîte de dialogue du Store Outlook 2016 avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. <span data-ttu-id="8ee69-142">Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue.</span><span class="sxs-lookup"><span data-stu-id="8ee69-142">Locate the **Custom add-ins** section at the bottom of the dialog.</span></span> <span data-ttu-id="8ee69-143">Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.</span><span class="sxs-lookup"><span data-stu-id="8ee69-143">Select the **Add a custom add-in** link, and then select **Add from file**.</span></span>

    ![Capture d’écran de la page Store avec une flèche pointant vers l’option À partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. <span data-ttu-id="8ee69-p106">Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.</span><span class="sxs-lookup"><span data-stu-id="8ee69-p106">Locate the manifest file for your custom add-in and install it. Accept all prompts during the installation.</span></span>

## <a name="remove-a-sideloaded-add-in"></a><span data-ttu-id="8ee69-147">Supprimer un complément versions test chargées</span><span class="sxs-lookup"><span data-stu-id="8ee69-147">Remove a sideloaded add-in</span></span>

<span data-ttu-id="8ee69-148">Pour supprimer un complément versions test chargées à partir d’Outlook, suivez les étapes décrites précédemment dans cet article pour trouver le complément dans la section **compléments personnalisés** de la boîte de dialogue qui répertorie vos compléments installés. Choisissez les points de suspension (`...`) pour le complément, puis cliquez sur **supprimer** pour supprimer ce complément spécifique.</span><span class="sxs-lookup"><span data-stu-id="8ee69-148">To remove a sideloaded add-in from Outlook, use the steps previously described in this article to find the add-in in the **Custom add-ins** section of the dialog box that lists your installed add-ins. Choose the ellipsis (`...`) for the the add-in and then choose **Remove** to remove that specific add-in.</span></span>