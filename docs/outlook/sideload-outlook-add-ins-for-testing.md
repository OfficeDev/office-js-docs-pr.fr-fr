---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 03/17/2022
ms.localizationpriority: medium
ms.openlocfilehash: 5f9a76aec20574b5397df9fcfd18e4714e5c8dd4
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745717"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>Chargement de version test des compléments Outlook

Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.

> [!IMPORTANT]
> Si votre Outlook prend en charge les appareils mobiles, chargez une version test du manifeste à l’aide des instructions de cet article pour votre client Outlook sur le web, Windows ou Mac, puis suivez les **instructions** de la section Test de vos macros supplémentaires sur mobile de l’article Sur les [macros pour Outlook Mobile](outlook-mobile-addins.md#testing-your-add-ins-on-mobile).

## <a name="sideload-automatically"></a>Chargement de version de version automatique

Si vous avez créé votre Outlook à l’aide du générateur [Yeoman](../develop/yeoman-generator-overview.md) pour les Office, il est préférable de recharger une version de version par le biais de la ligne de commande Windows. Cela tirera parti de nos outils et chargements de version de version sur tous vos appareils pris en charge dans une seule commande.

1. Sur Windows, ouvrez une invite de commandes et accédez au répertoire racine de votre projet de add-in généré par Yeoman. Exécutez la commande `npm start`.

1. Votre Outlook de bureau est automatiquement chargé de manière Outlook sur votre ordinateur de bureau. Une boîte de dialogue s’affiche, indiquant qu’il y a une tentative de chargement de version de chargement du module, répertoriant le nom et l’emplacement du fichier manifeste. **Sélectionnez OK**, qui enregistre le manifeste.

    > [!IMPORTANT]
    > Si le manifeste contient une erreur ou si le chemin d’accès au manifeste n’est pas valide, vous recevrez un message d’erreur.

1. Si votre manifeste ne contient aucune erreur et que le chemin d’accès est valide, votre application est désormais rechargée de côté et disponible à la fois sur votre bureau et dans Outlook sur le web. Il sera également installé sur tous vos appareils pris en charge.

## <a name="sideload-manually"></a>Chargement de version de version manuelle

Bien que nous recommandions vivement le chargement d’une version de version secondaire automatiquement par le biais de la ligne de commande comme abordé dans la section précédente, vous pouvez également charger manuellement une version de version de chargement de version de version de base d’un Outlook basé sur le client Outlook.

### <a name="outlook-on-the-web"></a>Outlook sur le web

Le processus de chargement d’une version de version Outlook sur le web dépend de l’utilisation de la nouvelle version ou de la version classique.

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#new-outlook-on-the-web).

    ![Capture d’écran partielle de la nouvelle barre Outlook sur le web’outils.](../images/outlook-on-the-web-new-toolbar.png)

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#classic-outlook-on-the-web).

    ![Capture d’écran partielle de la barre d’outils Outlook sur le web classique.](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.

#### <a name="new-outlook-on-the-web"></a>Nouvelle Outlook sur le web

1. Accédez à [Outlook sur le web](https://outlook.office.com).

1. Créez un message.

1. Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.

    ![Fenêtre de composition de message dans la nouvelle Outlook sur le web avec l’option Obtenir des add-ins mise en évidence.](../images/outlook-on-the-web-new-get-add-ins.png)

1. Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.

    ![Les applications pour Outlook boîte de dialogue dans la nouvelle Outlook sur le web avec Mes applications sélectionnées.](../images/outlook-on-the-web-new-my-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran gérer les add-ins pointant vers Ajouter à partir d’une option de fichier.](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

#### <a name="classic-outlook-on-the-web"></a>Modèle Outlook sur le web

1. Accédez à [Outlook sur le web](https://outlook.office.com).

1. Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.

    ![Outlook sur le web capture d’écran pointant vers l’option Gérer les add-ins.](../images/outlook-sideload-web-manage-integrations.png)

1. Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.

    ![Outlook sur le web dans la boîte de dialogue du Store avec Mes applications sélectionnée.](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran gérer les add-ins pointant vers Ajouter à partir d’une option de fichier.](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

### <a name="outlook-on-the-desktop"></a>Outlook sur le bureau

#### <a name="outlook-2016-or-later-on-windows-or-mac"></a>Outlook 2016 ou ultérieur sur Windows ou Mac

1. Ouvrez Outlook 2016 ou une ultérieure sur Windows mac.

1. Cliquez sur le bouton **Obtenir des compléments** du ruban.

    # <a name="windows"></a>[Fenêtres](#tab/windows)

    ![Outlook sur Windows ruban pointant vers le bouton Obtenir des modules.](../images/outlook-sideload-desktop-windows.png)

    # <a name="mac"></a>[Mac](#tab/mac)

    ![Outlook ruban Mac pointant vers le bouton Obtenir des macros.](../images/outlook-sideload-mac-classic.png)

    ---

    > [!IMPORTANT]
    > Si vous ne voyez pas le bouton Obtenir des Outlook, faites l’une des **choses** suivantes :
    >
    > # <a name="windows"></a>[Fenêtres](#tab/windows)
    >
    > - Si vous avez configuré la disposition du ruban sur Ruban simplifié, sélectionnez le bouton de sélection (`...`) dans le ruban, puis sélectionnez Obtenir **des modules.**
    >
    >   ![Outlook sur Windows pointant vers le bouton Obtenir des modules à partir du bouton de points de suite.](../images/outlook-sideload-simplified-ribbon.png)
    >
    > - Sélectionnez **le bouton Store** sur le ruban, si disponible.
    >
    > - Sélectionnez **le** menu Fichier, puis sélectionnez le bouton Gérer les **modules complémentaires** sous l’onglet  **Informations** pour ouvrir la boîte de dialogue Des Outlook sur le web. Vous pouvez en savoir plus sur l’expérience web dans la section précédente chargement de version de chargement d’un [Outlook sur le web](#outlook-on-the-web).
    >
    > # <a name="mac"></a>[Mac](#tab/mac)
    >
    > - Sélectionnez le bouton de sélection (`...`) dans le ruban, puis **sélectionnez Obtenir des modules.**
    >
    >   ![Outlook mac pointant vers le bouton Obtenir des macros à partir du bouton de points de suite.](../images/outlook-sideload-deskop-mac.png)
    >
    > - Sélectionnez **le bouton Store** sur le ruban, si disponible.
    >
    > ---

1. S’il existe des onglets en haut de la boîte de dialogue, **assurez-vous** que l’onglet Des applications est sélectionné. Choose **My add-ins**.

    ![Outlook 2016 dans la boîte de dialogue Du store avec Mes applications sélectionnée.](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran du magasin pointant sur Ajouter à partir d’une option de fichier.](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

#### <a name="outlook-2013-on-windows"></a>Outlook 2013 sous Windows

1. Ouvrez Outlook 2013 sur Windows.

1. Sélectionnez **le** menu Fichier, puis sélectionnez le bouton Gérer les **modules complémentaires** sous l’onglet **Informations**. Outlook ouvre la version web dans un navigateur.

1. Suivez les étapes de la section [Chargement de version de version](#outlook-on-the-web) Outlook sur le web en fonction de votre version de Outlook sur le web.

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un add-in chargé de nouveau

Sur toutes les versions de Outlook, la clé de la suppression d’un module de chargement de version ultérieure est  la boîte de dialogue Mes applications qui répertorie vos applications installées. Choisissez les ellipses (`...`) du module, puis sélectionnez **Supprimer**.

Pour accéder à la boîte de dialogue Mes applications pour votre client Outlook, utilisez les dernières **étapes** répertoriées pour le chargement de version [](#sideload-manually) manuelle dans les sections précédentes de cet article.

Pour supprimer un add-in chargé de Outlook, utilisez les étapes décrites précédemment dans cet article pour rechercher le module dans la section Custom **add-ins** de la boîte de dialogue qui répertorie vos applications installées. Choisissez les ellipses (`...`) du module, puis choisissez **Supprimer** pour supprimer ce dernier. Fermez la boîte de dialogue.

## <a name="see-also"></a>Voir aussi

- [Compléments pour Outlook Mobile](outlook-mobile-addins.md)