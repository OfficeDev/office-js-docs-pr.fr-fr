---
title: Chargement indépendant des compléments Office sur iPad à des fins de test
description: Testez votre complément Office sur iPad en effectuant un chargement indépendant.
ms.date: 06/29/2022
ms.localizationpriority: medium
ms.openlocfilehash: 0ba52ae78bed36c4eb8130c714577a1b0899aeb6
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713200"
---
# <a name="sideload-office-add-ins-on-ipad-for-testing"></a>Chargement indépendant des compléments Office sur iPad à des fins de test

Pour voir comment votre complément s’exécutera dans Office sur iOS, vous pouvez charger le manifeste de votre complément sur un iPad à l’aide d’iTunes. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.

> [!NOTE]
> Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="prerequisites-for-office-on-ios"></a>Configuration requise pour Office sur iOS

- Un ordinateur Windows ou Mac sur lequel [iTunes](https://www.apple.com/itunes/download/) est installé.
  > [!IMPORTANT]
  > Si vous exécutez macOS Catalina, [iTunes n’est plus disponible](https://support.apple.com/HT210200) . Vous devez donc suivre les instructions de la section [Charger un complément sur Excel ou Word sur iPad à l’aide de macOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) plus loin dans cet article.

- Un iPad exécutant iOS 8.2 ou version ultérieure avec [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) ou [Word](https://apps.apple.com/app/microsoft-word/id586447913) installé, et un câble de synchronisation.

- Le fichier .xml de manifeste pour le complément que vous voulez tester.

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>Charger une version test d’un complément sur Excel ou Word sur iPad à l’aide d’iTunes

1. Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Si vous connectez l’iPad à votre ordinateur pour la première fois, vous êtes invité à approuver **cet ordinateur.** Choisissez **Confiance** pour continuer.

2. Dans iTunes, sélectionnez l’icône **iPad** en dessous de la barre de menu.

3. Sous **Réglages** sur le côté gauche d’iTunes, sélectionnez **Applications**.

4. Sur le côté droite d’iTunes, faites défiler vers **Partage de fichiers**, puis sélectionnez **Excel** ou **Word** dans la colonne **Compléments**.

5. En bas de la colonne **Documents Excel** ou **Word** , choisissez **Ajouter un fichier**, puis sélectionnez le manifeste .xml fichier du complément que vous souhaitez charger de manière indépendante.

6. Ouvrez l'application Excel ou Word sur votre iPad. Si l’application Excel ou Word est déjà en cours d’exécution, choisissez le bouton **Accueil** , puis fermez et redémarrez l’application.

7. Ouvrez un document.

8. Sélectionnez **Compléments sous** l’onglet **Insertion** . (Sous l’onglet **Insertion** , vous devrez **peut-être** faire défiler horizontalement jusqu’à ce que le bouton Compléments s’affiche.) Votre complément chargé de manière indépendante est disponible à insérer sous l’en-tête **Développeur** dans l’interface utilisateur **des compléments** .

    ![Insérez des compléments dans l’application Excel.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>Charger une version test d’un complément sur Excel ou Word sur iPad à l’aide de macOS Catalina

> [!IMPORTANT]
> Avec l’introduction de macOS Catalina, [Apple a abandonné iTunes sur Mac](https://support.apple.com/HT210200) et les fonctionnalités intégrées requises pour charger des applications dans **Finder**.

1. Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Si vous connectez l’iPad à votre ordinateur pour la première fois, vous êtes invité à approuver **cet ordinateur.** Choisissez **Confiance** pour continuer. Vous pouvez également être invité à indiquer s’il s’agit d’un nouvel iPad ou si vous en restaurez un.

2. Dans Finder, sous **Emplacements**, choisissez l’icône **iPad** sous la barre de menus.

3. En haut de la fenêtre Finder, cliquez sur **Fichiers**, puis recherchez **Excel** ou **Word**.

4. À partir d’une autre fenêtre finder, faites glisser et déposez le fichier manifest.xml du complément que vous souhaitez charger de côté sur le fichier **Excel** ou **Word** dans la première fenêtre Finder.

5. Ouvrez l'application Excel ou Word sur votre iPad. Si l’application Excel ou Word est déjà en cours d’exécution, choisissez le bouton **Accueil** , puis fermez et redémarrez l’application.

6. Ouvrez un document.

7. Sélectionnez **Compléments sous** l’onglet **Insertion** . (Sous l’onglet **Insertion** , vous devrez **peut-être** faire défiler horizontalement jusqu’à ce que le bouton Compléments s’affiche.) Votre complément chargé de manière indépendante est disponible à insérer sous l’en-tête **Développeur** dans l’interface utilisateur **des compléments** .

    ![Insérez des compléments dans l’application Excel.](../images/excel-insert-add-in.png)

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément chargé de manière indépendante

Vous pouvez supprimer un complément précédemment chargé en désactivant le cache Office sur votre ordinateur. Pour plus d’informations sur l’effacement du cache pour chaque plateforme et application, consultez l’article [Effacer le cache Office](clear-cache.md).

## <a name="see-also"></a>Voir aussi

- [Chargement indépendant des compléments Office sur Mac à des fins de test](sideload-an-office-add-in-on-mac.md)
- [Déboguer des compléments Office sur un Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook pour les tester](../outlook/sideload-outlook-add-ins-for-testing.md)
