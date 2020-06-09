---
title: Chargement de version test des compléments Outlook
description: Utilisez le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.
ms.date: 06/24/2019
localization_priority: Normal
ms.openlocfilehash: 3543eeb58f441819edb2c129e6e14206e26de524
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44605324"
---
# <a name="sideload-outlook-add-ins-for-testing"></a>Chargement de version test des compléments Outlook

Vous pouvez utiliser le chargement de version test pour installer un complément Outlook sans avoir à le placer au préalable dans un catalogue de compléments.


## <a name="sideload-an-add-in-in-outlook-in-office-365"></a>Chargement d’une version test d’un complément dans Outlook dans Office 365

Le processus de chargement de la version test d’un complément dans Outlook dans Office 365 dépend de si vous utilisez la nouvelle version d’Outlook sur le web ou la version classique.

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

1. Ouvrez Outlook 2013 ou une version ultérieure sur Windows, ou Outlook 2016 ou une version ultérieure sur Mac.

1. Cliquez sur le bouton **Obtenir des compléments** du ruban.

    ![Ruban Outlook 2016 avec une flèche pointant sur le bouton Store](../images/outlook-sideload-desktop-store.png)

    > [!NOTE]
    > Si vous ne voyez pas le bouton **Obtenir des compléments** dans votre version d’Outlook, cliquez sur le bouton **Store** situé dans le ruban à la place.

1. Sélectionnez **Compléments**, puis **Mes compléments**.

    ![Boîte de dialogue du Store Outlook 2016 avec Mes compléments sélectionné](../images/outlook-sideload-store-select-add-ins.png)

1. Recherchez la section **Compléments personnalisés** en bas de la boîte de dialogue. Sélectionnez le lien **Ajouter un complément personnalisé**, puis sélectionnez **Ajouter à partir d’un fichier**.

    ![Capture d’écran de la page Store avec une flèche pointant vers l’option À partir d’un fichier](../images/outlook-sideload-desktop-add-from-file.png)

1. Localisez le fichier manifeste de votre complément personnalisé et installez-le. Acceptez toutes les invites pendant l’installation.

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément versions test chargées

Pour supprimer un complément versions test chargées à partir d’Outlook, suivez les étapes décrites précédemment dans cet article pour trouver le complément dans la section **compléments personnalisés** de la boîte de dialogue qui répertorie vos compléments installés. Choisissez les points de suspension ( `...` ) pour le complément, puis cliquez sur **supprimer** pour supprimer ce complément spécifique.