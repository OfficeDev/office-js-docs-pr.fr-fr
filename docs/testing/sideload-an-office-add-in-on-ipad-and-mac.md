---
title: Chargement de version test des compléments Office sur iPad et Mac
description: Testez votre Office sur iPad Mac en chargeant une version test.
ms.date: 09/24/2021
ms.localizationpriority: medium
ms.openlocfilehash: b57b072df1fa7c55e709f4ed4045cece8b95aa7e
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746611"
---
# <a name="sideload-office-add-ins-on-ipad-and-mac-for-testing"></a>Chargement de version test des compléments Office sur iPad et Mac

Pour voir comment votre complément s’exécutera dans Office sur iOS, vous pouvez charger une version test du manifeste de votre complément sur un iPad à l’aide d’iTunes ou directement dans Office sur Mac. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.

> [!NOTE]
> Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="prerequisites-for-office-on-ios"></a>Configuration requise pour Office sur iOS

- Un ordinateur Windows ou Mac sur lequel [iTunes](https://www.apple.com/itunes/download/) est installé.
  > [!IMPORTANT]
  > Si vous exécutez macOS Journal, [iTunes](https://support.apple.com/HT210200) n’est plus disponible. Vous devez donc suivre les instructions de la section Chargement de version de version ultérieure d’un [add-in sur Excel ou Word sur iPad à l’aide de macOS Importez plus](#sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina) loin dans cet article.

- Un iPad exécutant iOS 8.2 ou ultérieur avec Excel [ou Word](https://apps.apple.com/app/microsoft-word/id586447913) installé, [](https://apps.apple.com/app/microsoft-excel/id586683407) et un câble de synchronisation.

- Le fichier .xml de manifeste pour le complément que vous voulez tester.

## <a name="prerequisites-for-office-on-mac"></a>Configuration requise pour Office sur Mac

- Un Mac fonctionnant sous OS X v10.10 « Yosemite » ou une version ultérieure, avec [Office sur Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installé.

- Word sur Mac version 15.18 (160109).

- Excel sur Mac version 15.19 (160206).

- PowerPoint sur Mac version 15.24 (160614)

- Le fichier .xml de manifeste pour le complément que vous voulez tester.

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-itunes"></a>Chargement d’une version de version Excel ou Word sur iPad à l’aide d’iTunes

1. Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Si vous connectez le iPad à votre ordinateur pour la première fois, vous serez invité à nous faire confiance **à cet ordinateur**. **Sélectionnez Confiance** pour continuer.

2. Dans iTunes, sélectionnez l’icône **iPad** en dessous de la barre de menu.

3. Sous **Réglages** sur le côté gauche d’iTunes, sélectionnez **Applications**.

4. Sur le côté droite d’iTunes, faites défiler vers **Partage de fichiers**, puis sélectionnez **Excel** ou **Word** dans la colonne **Compléments**.

5. En bas de la colonne **Excel** ou **Documents Word**, sélectionnez Ajouter un **fichier, puis** sélectionnez le fichier .xml manifeste du module de chargement de version de version.

6. Ouvrez l'application Excel ou Word sur votre iPad. Si l Excel ou l’application Word est déjà en cours d’exécution, sélectionnez le bouton Accueil, puis fermez et redémarrez l’application.

7. Ouvrez un document.

8. Sélectionnez **Les add-ins** sous l’onglet Insérer.  (Sous l’onglet Insertion, vous devrez **peut-être** faire défiler horizontalement jusqu’à ce que vous voyez le bouton De plus.) Votre version de chargement de version de votre application peut être insérée  sous l’en-tête Développeur dans **l’interface utilisateur des applications**.

    ![Insérez des applications dans l Excel appl.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-on-excel-or-word-on-ipad-using-macos-catalina"></a>Chargement d’une version de version Excel ou Word sur iPad à l’aide de macOS

> [!IMPORTANT]
> Avec l’introduction de macOS Android, Apple a abandonné [iTunes sur Mac](https://support.apple.com/HT210200) et les fonctionnalités intégrées requises pour télécharger une version de version de chargement d’applications dans **Finder**.

1. Utilisez un câble de synchronisation pour connecter votre iPad à votre ordinateur. Si vous connectez le iPad à votre ordinateur pour la première fois, vous serez invité à nous faire confiance **à cet ordinateur**. **Sélectionnez Confiance** pour continuer. Vous pouvez également être invité à savoir s’il s’agit d’iPad ou si vous en restétiez un.

2. Dans Rechercher, sous **Emplacements**, sélectionnez **l’icône iPad** sous la barre de menus.

3. En haut de la fenêtre Finder, cliquez sur **Fichiers**, puis recherchez **Excel** ou **Word**.

4. Dans une autre fenêtre Finder, faites glisser et déposez le fichier manifest.xml du module que vous souhaitez charger de manière latérale sur le **fichier Excel** ou **Word** dans la première fenêtre Finder.

5. Ouvrez l'application Excel ou Word sur votre iPad. Si l Excel ou l’application Word est déjà en cours d’exécution, sélectionnez le bouton Accueil, puis fermez et redémarrez l’application.

6. Ouvrez un document.

7. Sélectionnez **Les add-ins** sous l’onglet Insérer.  (Sous l’onglet Insertion, vous devrez **peut-être** faire défiler horizontalement jusqu’à ce que vous voyez le bouton De plus.) Votre version de chargement de version de votre application peut être insérée  sous l’en-tête Développeur dans **l’interface utilisateur des applications**.

    ![Insérez des applications dans l Excel appl.](../images/excel-insert-add-in.png)

## <a name="sideload-an-add-in-in-office-on-mac"></a>Chargement d’une version test de complément dans Office sur Mac

1. **Ouvrez Terminal** et allez dans l’un des dossiers suivants où vous allez enregistrer le fichier manifeste de votre module. Si le dossier `wef` n’existe pas sur votre ordinateur, créez-le.

    - Pour Word : `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Pour Excel : `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Pour PowerPoint : `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

2. Ouvrez le dossier dans **Finder à** l’aide de la commande `open .` (y compris le point ou le point). Copier le fichier manifeste de votre complément dans ce dossier.

    ![Dossier Wef dans Office sur Mac.](../images/all-my-files.png)

3. Ouvrez Word, puis ouvrez un document. Redémarrez Word si cette application est déjà en cours d'exécution.

4. Dans Word, choisissez **InsertAdd-insMy** >  >  **Add-ins** (menu déroulant), puis choisissez votre add-in.

    ![Mes Office sur Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Les versions test chargées de vos compléments ne s’afficheront pas dans la boîte de dialogue Mes compléments. Elles sont visibles uniquement dans le menu déroulant (petite flèche vers le bas à droite de Mes compléments dans l’onglet **Insérer**). Les versions test chargées de vos compléments sont répertoriées sous l’en-tête **Compléments de développeur** dans ce menu.

5. Vérifiez que votre complément apparaît dans Word.

    ![Office’affichage d’un Office sur Mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un add-in chargé de nouveau

Vous pouvez supprimer un add-in précédemment chargé de nouveau en effasant le cache Office sur votre ordinateur. Pour plus d’informations sur la façon de effacer le cache pour chaque plateforme et application, voir l’article Effacer [le cache Office cache](clear-cache.md).

## <a name="see-also"></a>Voir aussi

- [Déboguer des compléments Office sur un Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md)
