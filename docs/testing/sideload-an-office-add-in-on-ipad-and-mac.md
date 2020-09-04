---
title: Chargement de version test des compléments Office sur iPad et Mac
description: Testez votre complément Office sur iPad et Mac par chargement.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 7c5e9542c6e6f9abc96defde389b9543421b8529
ms.sourcegitcommit: 604361e55dee45c7a5d34c2fa6937693c154fc24
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/03/2020
ms.locfileid: "47364054"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Chargement de version test des compléments Office sur iPad et Mac

Pour voir comment votre complément s’exécutera dans Office sur iOS, vous pouvez charger une version test du manifeste de votre complément sur un iPad à l’aide d’iTunes ou directement dans Office sur Mac. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.

## <a name="prerequisites-for-office-on-ios"></a>Configuration requise pour Office sur iOS

- Un ordinateur Windows ou Mac sur lequel [iTunes](https://www.apple.com/itunes/download/) est installé.
  > [!IMPORTANT]
  > Si vous exécutez macOS Catalina, [iTunes n’est plus disponible](https://support.apple.com/HT210200) et vous devez suivre les instructions de la section [chargement d’un complément sur Excel ou de Word sur iPad à l’aide de MacOS Catalina](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) plus loin dans cet article.

- Un iPad exécutant iOS 8,2 ou version ultérieure avec [Excel](https://apps.apple.com/app/microsoft-excel/id586683407) ou [Word](https://apps.apple.com/app/microsoft-word/id586447913) et un câble de synchronisation.

- Le fichier .xml de manifeste pour le complément que vous voulez tester.

## <a name="prerequisites-for-office-on-mac"></a>Configuration requise pour Office sur Mac

- Un Mac fonctionnant sous OS X v10.10 « Yosemite » ou une version ultérieure, avec [Office sur Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installé.

- Word sur Mac version 15.18 (160109).

- Excel sur Mac version 15.19 (160206).

- PowerPoint sur Mac version 15.24 (160614)

- Le fichier .xml de manifeste pour le complément que vous voulez tester.

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>Chargement d’un complément dans Excel ou Word sur iPad à l’aide d’iTunes

1. Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Si vous connectez l’ordinateur iPad à votre ordinateur pour la première fois, vous êtes invité à **approuver cet ordinateur ?**. Sélectionnez **approuver** pour continuer.

2. Dans iTunes, sélectionnez l’icône **iPad** en dessous de la barre de menu.

3. Sous **Réglages** sur le côté gauche d’iTunes, sélectionnez **Applications**.

4. Sur le côté droite d’iTunes, faites défiler vers **Partage de fichiers**, puis sélectionnez **Excel** ou **Word** dans la colonne **Compléments**.

5. Au bas de la colonne **Excel** ou **documents Word** , sélectionnez **Ajouter un fichier**, puis sélectionnez le fichier manifest. xml du complément que vous souhaitez chargement.

6. Ouvrez l'application Excel ou Word sur votre iPad. Si l’application Excel ou Word est déjà en cours d’exécution, cliquez sur le bouton **Accueil** , puis fermez et redémarrez l’application.

7. Ouvrez un document.

8. Choisissez **compléments** sous l’onglet **insertion** . (sous l’onglet **insertion** , vous devrez peut-être faire défiler horizontalement jusqu’à ce que le bouton **compléments** s’affiche.) Votre complément versions test chargées peut être inséré sous le titre **développeur** dans l’interface utilisateur des **compléments** .

    ![Insérer des compléments dans l’application Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>Chargement d’un complément sur Excel ou Word sur iPad à l’aide de macOS Catalina

> [!IMPORTANT]
> Avec l’introduction de macOS Catalina, [Apple a abandonné iTunes sur Mac](https://support.apple.com/HT210200) et les fonctionnalités intégrées requises pour chargement les applications dans **Finder**.

1. Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Si vous connectez l’ordinateur iPad à votre ordinateur pour la première fois, vous êtes invité à **approuver cet ordinateur ?**. Sélectionnez **approuver** pour continuer. Vous pouvez également être invité à indiquer s’il s’agit d’un nouvel iPad ou si vous effectuez une restauration.

2. Dans Finder, sous **emplacements**, sélectionnez l’icône **iPad** en dessous de la barre de menus.

3. En haut de la fenêtre Finder, cliquez sur **fichiers**, puis recherchez **Excel** ou **Word**.

4. Dans une fenêtre de recherche différente, glissez-déplacez le fichier de manifest.xml du complément à charger vers le fichier **Excel** ou **Word** dans la première fenêtre de recherche.

5. Ouvrez l'application Excel ou Word sur votre iPad. Si l’application Excel ou Word est déjà en cours d’exécution, cliquez sur le bouton **Accueil** , puis fermez et redémarrez l’application.

6. Ouvrez un document.

7. Choisissez **compléments** sous l’onglet **insertion** . (sous l’onglet **insertion** , vous devrez peut-être faire défiler horizontalement jusqu’à ce que le bouton **compléments** s’affiche.) Votre complément versions test chargées peut être inséré sous le titre **développeur** dans l’interface utilisateur des **compléments** .

    ![Insérer des compléments dans l’application Excel](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Chargement d’une version test de complément dans Office sur Mac

> [!NOTE]
> Pour charger une version test de complément Outlook sur Mac, voir l’article relatif au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-an-add-in-in-outlook-on-the-desktop).

1. Ouvrez le **Terminal** et accédez à l’un des dossiers suivants, dans lequel vous allez enregistrer le fichier manifeste de votre complément. Si le dossier `wef` n’existe pas sur votre ordinateur, créez-le.

    - Pour Word : `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Pour Excel : `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Pour PowerPoint : `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. Ouvrez le dossier dans **Finder** à l’aide de la commande `open .` (y compris le point ou le point). Copier le fichier manifeste de votre complément dans ce dossier.

    ![Dossier WEF dans Office sur Mac](../images/all-my-files.png)

3. Ouvrez Word, puis ouvrez un document. Redémarrez Word si cette application est déjà en cours d'exécution.

4. Dans Word, sélectionnez **Insérer**des  >  **compléments**  >  **My Add-ins** (menu déroulant), puis choisissez votre complément.

    ![Mes compléments dans Office sur Mac](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Les versions test chargées de vos compléments ne s’afficheront pas dans la boîte de dialogue Mes compléments. Elles sont visibles uniquement dans le menu déroulant (petite flèche vers le bas à droite de Mes compléments dans l’onglet **Insérer**). Les versions test chargées de vos compléments sont répertoriées sous l’en-tête **Compléments de développeur** dans ce menu.

5. Vérifiez que votre complément apparaît dans Word.

    ![Complément Office affiché dans Office sur Mac](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément versions test chargées

Vous pouvez supprimer un complément précédemment versions test chargées en effaçant le cache Office sur votre ordinateur. Pour plus d’informations sur la façon d’effacer le cache de chaque plateforme et application, consultez l’article [effacer le cache Office](clear-cache.md).

## <a name="see-also"></a>Voir aussi

- [Débogage des compléments Office sur iPad et Mac](debug-office-add-ins-on-ipad-and-mac.md)
