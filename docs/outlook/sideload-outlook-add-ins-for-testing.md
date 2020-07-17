---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 07/09/2020
localization_priority: Normal
ms.openlocfilehash: 9b44b988ddd6552d5f7d14088a0b6f3ae1e410ed
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093881"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>Chargement de version test des compléments Outlook

Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.

## <a name="sideload-an-add-in-in-outlook-on-the-web"></a>Chargement d’un complément dans Outlook sur le web

Le processus de chargement d’un complément dans Outlook sur le Web dépend de si vous utilisez la version nouvelle ou classique.

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la nouvelle version d’Outlook sur le web](#sideload-an-add-in-in-the-new-outlook-on-the-web).

    ![capture d’écran partielle de la nouvelle version de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-new-toolbar.png)

- Si la barre d’outils de boîte aux lettres ressemble à l’image suivante, reportez-vous à la section relative au [chargement de la version test d’un complément dans la version classique d’Outlook sur le web](#sideload-an-add-in-in-classic-outlook-on-the-web).

    ![capture d’écran partielle de la version classique de la barre d’outils d’Outlook sur le web](../images/outlook-on-the-web-classic-toolbar.png)

> [!NOTE]
> Si votre organisation a inclus son logo dans la barre d’outils de boîte aux lettres, le rendu sera peut-être légèrement différent de celui figurant dans les images précédentes.

### <a name="sideload-an-add-in-in-the-new-outlook-on-the-web"></a>Chargement d’un complément dans la nouvelle version d’Outlook sur le web

1. Accédez à [Outlook dans Office 365](https://outlook.office.com).

1. Dans Outlook sur le web, créez un message.

1. Sélectionnez **...** au bas du nouveau message, puis sélectionnez **Obtenir des compléments** dans le menu qui s’affiche.

    ![Fenêtre de composition de messages dans la nouvelle version d’Outlook sur le web avec l’option pour obtenir des compléments en évidence](../images/outlook-on-the-web-new-get-add-ins.png)

1. Dans la boîte de dialogue **Compléments pour Outlook**, sélectionnez **Mes compléments**.

    ![Boîte de dialogue Compléments pour Outlook dans la nouvelle version d’Outlook sur le web avec l’option Mes compléments sélectionnée](../images/outlook-on-the-web-new-my-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

### <a name="sideload-an-add-in-in-classic-outlook-on-the-web"></a>Chargement d’un complément dans la version classique d’Outlook sur le web

1. Accédez à [Outlook dans Office 365](https://outlook.office.com).

1. Cliquez sur l’icône en forme d’engrenage située en haut à droite de la barre d’outils et sélectionnez **Gérer des compléments**.

    ![Capture d’écran d’Outlook sur le web avec une flèche pointant sur l’option Gérer les compléments](../images/outlook-sideload-web-manage-integrations.png)

1. Sur la page **Gérer les compléments**, sélectionnez **Compléments**, puis **Mes compléments**.

    ![Boîte de dialogue du Store Outlook sur le web avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de gestion des compléments pointant vers l’option Ajouter à partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

## <a name="sideload-an-add-in-in-outlook-on-the-desktop"></a>Chargement d’un complément dans la version de bureau d’Outlook

### <a name="outlook-2016-or-later"></a>Outlook 2016 ou version ultérieure

1. Ouvrez Outlook 2016 ou une version ultérieure sur Windows ou Mac.

1. Cliquez sur le bouton **Obtenir des compléments** du ruban.

    ![Ruban Outlook 2016 avec une flèche pointant sur le bouton Store](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > Si vous ne voyez pas le bouton **Obtenir des compléments** dans votre version d’Outlook, cliquez sur le bouton **Store** situé dans le ruban à la place.

1. Sélectionnez **Compléments**, puis **Mes compléments**.

    ![Boîte de dialogue du Store Outlook 2016 avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de la page Store avec une flèche pointant vers l’option À partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

### <a name="outlook-2013"></a>Outlook 2013

1. Ouvrez Outlook 2013 sur Windows.

1. Sélectionnez le menu **fichier** , puis cliquez sur le bouton **gérer les compléments** sous l’onglet **informations** . Outlook ouvre un navigateur.

1. Suivez les étapes de la section [chargement d’un complément dans Outlook sur le Web](#sideload-an-add-in-in-outlook-on-the-web) en fonction de votre version d’Outlook sur le Web.

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément versions test chargées

Pour supprimer un complément versions test chargées à partir d’Outlook, suivez les étapes décrites précédemment dans cet article pour trouver le complément dans la section **compléments personnalisés** de la boîte de dialogue qui répertorie vos compléments installés. Choisissez les points de suspension ( `...` ) pour le complément, puis cliquez sur **supprimer** pour supprimer ce complément spécifique.