---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 02/10/2021
localization_priority: Normal
ms.openlocfilehash: b783b815af84a7fd8b4abd52cdd8e0925bfb9ecf
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234246"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>Chargement de version test des compléments Outlook

Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.

## <a name="sideload-automatically"></a>Chargement de version de version automatique

Si vous avez créé votre add-in Outlook à l’aide du générateur Yeoman pour les [add-ins Office,](https://github.com/OfficeDev/generator-office)il est préférable d’utiliser le chargement de version secondaire via la ligne de commande. Cela tirera parti de nos outils et chargements de version de version sur tous vos appareils pris en charge dans une seule commande.

1. À l’aide de la ligne de commande, accédez au répertoire racine de votre projet de add-in généré par Yeoman. Exécutez la commande `npm start`.

2. Votre application Outlook charge automatiquement une version de version vers Outlook sur votre ordinateur de bureau. Une boîte de dialogue s’affiche, indiquant qu’il y a une tentative de chargement de version de chargement du module, répertoriant le nom et l’emplacement du fichier manifeste. Sélectionnez **OK,** qui enregistre le manifeste.

> [!IMPORTANT]
> Si le manifeste contient une erreur ou si le chemin d’accès au manifeste n’est pas valide, vous recevrez un message d’erreur.

3. Si votre manifeste ne contient aucune erreur et que le chemin d’accès est valide, votre application sera désormais rechargée de nouveau et disponible à la fois sur votre bureau et dans Outlook sur le web. Il sera également installé sur tous vos appareils pris en charge.

## <a name="sideload-manually"></a>Chargement de version de version manuelle

Bien que nous recommandions vivement le chargement de version secondaire automatiquement via la ligne de commande comme abordé dans la section précédente, vous pouvez également charger manuellement une version de version de chargement de version antérieure d’un add-in Outlook basé sur le client Outlook.

### <a name="outlook-on-the-web"></a>Outlook sur le web

Le processus de chargement d’une version d’évaluation d’un module dans Outlook sur le web varie selon que vous utilisez la nouvelle version ou la version classique.

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#new-outlook-on-the-web).

    ![capture d’écran partielle de la nouvelle version de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-new-toolbar.png)

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#classic-outlook-on-the-web).

    ![capture d’écran partielle de la version classique de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.

### <a name="new-outlook-on-the-web"></a>Nouvel Outlook sur le web

1. Accédez à [Outlook sur le web](https://outlook.office.com).

1. Créez un message.

1. Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.

    ![Fenêtre de composition de messages dans la nouvelle version d’Outlook sur le web avec l’option pour obtenir des compléments en évidence](../images/outlook-on-the-web-new-get-add-ins.png)

1. Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.

    ![Boîte de dialogue Compléments pour Outlook dans la nouvelle version d’Outlook sur le web avec l’option Mes compléments sélectionnée](../images/outlook-on-the-web-new-my-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

### <a name="classic-outlook-on-the-web"></a>Outlook sur le web classique

1. Accédez à [Outlook sur le web](https://outlook.office.com).

1. Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.

    ![Capture d’écran d’Outlook sur le web avec une flèche pointant sur l’option Gérer les compléments](../images/outlook-sideload-web-manage-integrations.png)

1. Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.

    ![Boîte de dialogue du Store Outlook sur le web avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

### <a name="outlook-on-the-desktop"></a>Outlook sur le bureau

#### <a name="outlook-2016-or-later"></a>Outlook 2016 ou une ultérieure

1. Ouvrez Outlook 2016 ou une édition ultérieure sur Windows ou Mac.

1. Cliquez sur le bouton **Obtenir des compléments** du ruban.

    ![Ruban Outlook 2016 pointant vers le bouton Obtenir des modules](../images/outlook-sideload-desktop-store.png)

    > [!IMPORTANT]
    > Si vous ne voyez pas le bouton Obtenir **des modules** dans votre version d’Outlook, sélectionnez :
    >
    > - **Bouton Stocker** sur le ruban, si disponible.
    >
    >   Ou
    >
    > - **Menu** Fichier, puis sélectionnez le bouton Gérer les **modules complémentaires** sous l’onglet **Informations** pour ouvrir la boîte de dialogue Des **modules complémentaires** dans Outlook sur le web.<br>Vous pouvez en savoir plus sur l’expérience web dans la section précédente Chargement d’une version de version de version antérieure d’un [add-in dans Outlook sur le web.](#outlook-on-the-web)

1. S’il existe des onglets en haut de la boîte de dialogue, **assurez-vous** que l’onglet Des applications est sélectionné. Choose **My add-ins**.

    ![Boîte de dialogue du Store Outlook 2016 avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de la page Store avec une flèche pointant vers l’option À partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

#### <a name="outlook-2013"></a>Outlook 2013

1. Ouvrez Outlook 2013 sur Windows.

1. Sélectionnez **le** menu Fichier, puis le bouton Gérer les **modules complémentaires** sous **l’onglet** Informations. Outlook ouvre la version web dans un navigateur.

1. Suivez les étapes de la section Chargement de version sideload d’un [add-in](#outlook-on-the-web) dans Outlook sur le web en fonction de votre version d’Outlook sur le web.

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un add-in chargé de nouveau

Sur toutes les versions d’Outlook, la clé de  la suppression d’un module de chargement secondaire est la boîte de dialogue Mes applications qui répertorie vos applications installées. Choisissez les ellipses ( ) pour le `...` add-in, puis sélectionnez **Supprimer**.

Pour accéder à la boîte de dialogue Mes applications pour votre client [](#sideload-manually) Outlook, utilisez les dernières **étapes** répertoriées pour le chargement de version manuelle dans les sections précédentes de cet article.

Pour supprimer un add-in chargé de côté d’Outlook, utilisez les étapes décrites précédemment dans cet article pour rechercher le module dans la section Custom **add-ins** de la boîte de dialogue qui répertorie vos applications installées. Choisissez les ellipses ( ) pour le module, puis choisissez Supprimer pour `...` supprimer ce dernier. 

