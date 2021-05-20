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
# <a name="sideload-outlook-add-ins-for-testing"></a>Chargement de version test des compléments Outlook

Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.

## <a name="sideload-automatically"></a>Sideload automatiquement

Si vous avez créé votre Outlook add-in en utilisant [le générateur Yeoman pour Office Add-ins,](https://github.com/OfficeDev/generator-office)sideloading est préférable de le faire à travers la ligne de commande. Cela profitera de notre outillage et de notre charge latérale sur tous vos appareils pris en charge en une seule commande.

1. À l’aide de la ligne de commande, accédez à l’annuaire racine de votre projet yeoman généré add-in. Exécutez la commande `npm start`.

1. Votre Outlook add-in sera automatiquement sideload pour Outlook votre ordinateur de bureau. Vous verrez apparaître un dialogue, indiquant qu’il y a une tentative de sideload l’add-in, énumérant le nom et l’emplacement du fichier manifeste. Sélectionnez **OK**, qui enregistrera le manifeste.

    > [!IMPORTANT]
    > Si le manifeste contient une erreur ou si le chemin vers le manifeste est invalide, vous recevrez un message d’erreur.

1. Si votre manifeste ne contient aucune erreur et que le chemin est valide, votre module sera désormais sideloaded et disponible à la fois sur votre bureau et dans les Outlook sur le Web. Il sera également installé sur tous vos appareils pris en charge.

## <a name="sideload-manually"></a>Sideload manuellement

Bien que nous vous recommandons fortement de recharger automatiquement à travers la ligne de commande telle que couverte dans la section précédente, vous pouvez également sideload manuellement un add-in Outlook basé sur le client Outlook.

### <a name="outlook-on-the-web"></a>Outlook sur le web

Le processus de chargement latéral d’un add-in Outlook sur le web dépend si vous utilisez la nouvelle version ou classique.

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#new-outlook-on-the-web).

    ![capture d’écran partielle de la nouvelle version de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-new-toolbar.png)

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#classic-outlook-on-the-web).

    ![capture d’écran partielle de la version classique de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.

### <a name="new-outlook-on-the-web"></a>Nouvelles Outlook sur le web

1. Accédez à [Outlook sur le web](https://outlook.office.com).

1. Créez un nouveau message.

1. Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.

    ![Fenêtre de composition de messages dans la nouvelle version d’Outlook sur le web avec l’option pour obtenir des compléments en évidence](../images/outlook-on-the-web-new-get-add-ins.png)

1. Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.

    ![Boîte de dialogue Compléments pour Outlook dans la nouvelle version d’Outlook sur le web avec l’option Mes compléments sélectionnée](../images/outlook-on-the-web-new-my-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

### <a name="classic-outlook-on-the-web"></a>Les Outlook classiques sur le web

1. Accédez à [Outlook sur le web](https://outlook.office.com).

1. Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.

    ![Capture d’écran d’Outlook sur le web avec une flèche pointant sur l’option Gérer les compléments](../images/outlook-sideload-web-manage-integrations.png)

1. Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.

    ![Boîte de dialogue du Store Outlook sur le web avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

### <a name="outlook-on-the-desktop"></a>Outlook sur le bureau

#### <a name="outlook-2016-or-later"></a>Outlook 2016 ou plus tard

1. Ouvrez Outlook 2016 ou plus tard sur Windows ou Mac.

1. Cliquez sur le bouton **Obtenir des compléments** du ruban.

    ![Outlook 2016 ruban pointant vers le bouton Get Add-ins](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > Si vous ne voyez pas le bouton **Get Add-ins** dans votre version de Outlook, sélectionnez :
    >
    > - **Rangez** le bouton sur le ruban, si disponible.
    >
    >   OR
    >
    > - **Menu** de fichiers, puis sélectionnez **le bouton Manage Add-ins** sur **l’onglet Info** pour ouvrir le dialogue **Add-ins** Outlook sur le web.<br>Vous pouvez en savoir plus sur l’expérience Web dans la section [précédente Sideload un add-in dans Outlook sur le web](#outlook-on-the-web).

1. S’il y a des onglets près du haut du dialogue, **assurez-vous que l’onglet Add-ins** est sélectionné. Choisissez **mes add-ins**.

    ![Boîte de dialogue du Store Outlook 2016 avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de la page Store avec une flèche pointant vers l’option À partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

#### <a name="outlook-2013"></a>Outlook 2013

1. Ouvert Outlook 2013 le Windows.

1. Sélectionnez **le** menu Fichier, puis **sélectionnez le bouton Manage Add-ins** sur l’onglet **Info.** Outlook ouvrira la version Web dans un navigateur.

1. Suivez les étapes du [Sideload un add-in dans Outlook sur la](#outlook-on-the-web) section web en fonction de votre version de Outlook sur le web.

## <a name="remove-a-sideloaded-add-in"></a>Retirer un add-in sideloaded

Sur toutes les versions de Outlook, la clé pour supprimer un module d’ajout sideloaded est le dialogue **My Add-ins** qui répertorie vos modules d’ajout installés. Choisissez l’ellipsis ( `...` ) pour l’add-in puis sélectionnez **Supprimer**.

Pour naviguer vers la **boîte de dialogue My Add-ins** pour votre client Outlook, utilisez les dernières étapes répertoriées pour le chargement manuel [dans](#sideload-manually) les sections précédentes de cet article.

Pour supprimer un module d’ajout sideloaded de Outlook, utilisez les étapes précédemment décrites dans cet article pour trouver l’add-in dans la section **add-ins personnalisés** de la boîte de dialogue qui répertorie vos modules supplémentaires installés. Choisissez l’ellipsis `...` ( ) pour l’add-in puis **choisissez Supprimer** pour supprimer cet add-in spécifique. Fermez la boîte de dialogue.
