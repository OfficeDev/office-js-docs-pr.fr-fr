---
title: Chargement indépendant des compléments Office sur Mac à des fins de test
description: Testez votre complément Office sur Mac en effectuant un chargement indépendant.
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 38ed5f5dba2d379b6137a098240021bd642d6e11
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/11/2022
ms.locfileid: "66713218"
---
# <a name="sideload-office-add-ins-on-mac-for-testing"></a>Chargement indépendant des compléments Office sur Mac à des fins de test

Pour voir comment votre complément s’exécutera sur Office sur Mac, vous pouvez charger de manière indépendante le manifeste de votre complément. Cette opération ne vous permettra pas de définir des points d’arrêt ni de déboguer le code de votre complément pendant son exécution, mais vous pourrez observer son comportement, et vérifier que l’interface utilisateur est fonctionnelle et qu’elle s’affiche correctement.

> [!NOTE]
> Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).

## <a name="prerequisites-for-office-on-mac"></a>Configuration requise pour Office sur Mac

- Un Mac fonctionnant sous OS X v10.10 « Yosemite » ou une version ultérieure, avec [Office sur Mac](https://products.office.com/buy/compare-microsoft-office-products?tab=omac) installé.

- Word sur Mac version 15.18 (160109).

- Excel sur Mac version 15.19 (160206).

- PowerPoint sur Mac version 15.24 (160614).

- Le fichier .xml de manifeste pour le complément que vous voulez tester.

## <a name="sideload-an-add-in-in-office-on-mac"></a>Chargement d’une version test de complément dans Office sur Mac

1. Utilisez **Finder** pour charger de manière indépendante le fichier manifeste. **Ouvrez Finder**, puis entrez Cmd+Maj+G pour ouvrir la boîte **de dialogue Accéder au dossier**.

1. Entrez l’un des chemins de fichiers suivants, en fonction de l’application que vous souhaitez utiliser pour le chargement indépendant. Si le dossier `wef` n’existe pas sur votre ordinateur, créez-le.

    - Pour Word : `/Users/<username>/Library/Containers/com.microsoft.Word/Data/Documents/wef`
    - Pour Excel : `/Users/<username>/Library/Containers/com.microsoft.Excel/Data/Documents/wef`
    - Pour PowerPoint : `/Users/<username>/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef`

        > [!NOTE]
        > Les étapes restantes décrivent comment charger une version test d’un complément Word.

1. Copiez le fichier manifeste de votre complément dans ce `wef` dossier.

    ![Dossier Wef dans Office sur Mac.](../images/all-my-files.png)

1. Ouvrez Word, puis ouvrez un document. Redémarrez Word si cette application est déjà en cours d'exécution.

1. Dans Word, choisissez **Insérer** > **des compléments** > **Mes compléments** (menu déroulant), puis choisissez votre complément.

    ![Mes compléments dans Office sur Mac.](../images/my-add-ins-wikipedia.png)

    > [!IMPORTANT]
    > Les versions test chargées de vos compléments ne s’afficheront pas dans la boîte de dialogue Mes compléments. Elles sont visibles uniquement dans le menu déroulant (petite flèche vers le bas à droite de Mes compléments dans l’onglet **Insérer**). Les versions test chargées de vos compléments sont répertoriées sous l’en-tête **Compléments de développeur** dans ce menu.

1. Vérifiez que votre complément apparaît dans Word.

    ![Complément Office affiché dans Office sur Mac.](../images/lorem-ipsum-wikipedia.png)

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément chargé de manière indépendante

Vous pouvez supprimer un complément précédemment chargé en désactivant le cache Office sur votre ordinateur. Pour plus d’informations sur l’effacement du cache pour chaque plateforme et application, consultez l’article [Effacer le cache Office](clear-cache.md).

## <a name="see-also"></a>Voir aussi

- [Chargement indépendant des compléments Office sur iPad à des fins de test](sideload-an-office-add-in-on-ipad.md)
- [Déboguer des compléments Office sur un Mac](debug-office-add-ins-on-ipad-and-mac.md)
- [Chargement de version test des compléments Outlook pour les tester](../outlook/sideload-outlook-add-ins-for-testing.md)
