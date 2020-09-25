---
title: Chargement de version test des compléments Office dans Office sur le web
description: Testez votre complément Office dans Office sur le Web par chargement.
ms.date: 09/24/2020
localization_priority: Normal
ms.openlocfilehash: 91f23200a2c393eb5c79f615765df52f205ac6e1
ms.sourcegitcommit: 09e1d8ff14b3c09a3eb11c91432c224a539181a4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/25/2020
ms.locfileid: "48268564"
---
# <a name="sideload-office-add-ins-in-office-on-the-web-for-testing"></a>Chargement de version test des compléments Office dans Office sur le web

Vous procéder à un chargement de version test pour installer un complément Office sans avoir à le placer au préalable dans un catalogue de compléments. Chargement peut être réalisé dans Microsoft 365 ou Office sur le Web. La procédure est légèrement différente entre les deux plateformes.

Lorsque vous chargez une version test d’un complément, le manifeste du complément est stocké dans le stockage local du navigateur. Ainsi, si vous videz le cache du navigateur ou si vous basculez vers un autre navigateur, vous devez à nouveau charger une version test de complément.

> [!NOTE]
> Tel que décrit dans cet article, le chargement de version test est pris en charge dans Word, Excel et PowerPoint. Pour charger une version test de complément Outlook, voir la rubrique relative au [chargement de version test des compléments Outlook](../outlook/sideload-outlook-add-ins-for-testing.md).

La vidéo suivante présente la procédure de chargement de version test de votre complément dans la version Office sur le web ou le bureau.

> [!VIDEO https://www.youtube.com/embed/XXsAw2UUiQo]

## <a name="sideload-an-office-add-in-in-office-on-the-web"></a>Chargement de version test d’un complément Office dans Office sur le web

1. Ouvrez [Office sur le Web](https://office.live.com/).

2. Dans **commencer à utiliser les applications en ligne maintenant**, choisissez **Excel**, **Word**ou **PowerPoint**; puis ouvrez un nouveau document.

3. Ouvrez l’onglet **Insérer** dans le ruban, puis dans la section **compléments** , choisissez **Compléments Office**.

4. Dans la boîte de dialogue **Compléments Office** , sélectionnez l’onglet **mes compléments** , choisissez **gérer mes compléments**, puis **Télécharger mon complément**.

    ![Boîte de dialogue Compléments Office avec une liste déroulante dans le coin supérieur droit indiquant « Gérer mes compléments » et une autre liste déroulante sous cette dernière avec l’option « Charger mon complément »](../images/office-add-ins-my-account.png)

5. **Accédez** au fichier manifeste du complément, puis sélectionnez **Télécharger**.

    ![Boîte de dialogue de téléchargement de complément avec des boutons pour parcourir, télécharger et annuler.](../images/upload-add-in.png)

6. Vérifiez que votre complément est installé. S’il s’agit d’une commande de complément, elle doit apparaître dans le ruban ou dans le menu contextuel. S’il s’agit d’un complément du volet Office, le volet doit apparaître.

> [!NOTE]
> Pour tester votre complément Office avec Microsoft Edge avec le WebView d’origine (EdgeHTML), une étape de configuration supplémentaire est requise. Dans une invite de commandes Windows, exécutez la commande suivante : `npx office-addin-dev-settings appcontainer EdgeWebView --loopback --yes` . Cela n’est pas obligatoire lorsque Office utilise le WebView2 Edge basé sur le chrome. Pour plus d’informations, consultez la rubrique [navigateurs utilisés par les compléments Office](../concepts/browsers-used-by-office-web-add-ins.md).

## <a name="sideload-an-office-add-in-in-office-365"></a>Chargement de version test d’un complément Office dans Office 365

1. Connectez-vous à votre compte Microsoft 365.

2. Ouvrez le lanceur d’applications à l’extrémité gauche de la barre d’outils et sélectionnez **Excel**, **Word**ou **PowerPoint**, puis créez un document.

3. Les étapes 3 à 6 sont identiques à celles de la section précédente, **Chargement d’une version de test d’un complément Office dans Office sur le web**.

## <a name="sideload-an-add-in-when-using-visual-studio"></a>Chargement d’une version test d’un complément lors de l’utilisation de Visual Studio

Si vous développez votre complément à l’aide de Visual Studio, le processus de chargement d’une version de teste est similaire. La seule différence est que vous devez mettre à jour la valeur de l’élément **SourceURL** dans votre manifeste afin d’inclure l’URL complète de déploiement du complément.

> [!NOTE]
> Si vous pouvez charger une version test des compléments à partir de Visual Studio vers Office sur le web, vous ne pouvez pas les déboguer à partir de Visual Studio. Pour déboguer, vous devrez utiliser les outils de débogage du navigateur. Pour plus d’informations, voir [Débogage de compléments dans Office sur le web](debug-add-ins-in-office-online.md).

1. Dans Visual Studio, affichez la fenêtre **Propriétés** en choisissant **Affichage** -> **Fenêtre Propriétés**.
2. Dans l’**Explorateur de solutions**, sélectionnez le projet web. Cela a pour effet d’afficher les propriétés du projet dans la fenêtre **Propriétés**.
3. Dans la fenêtre Propriétés, copiez l’**URL SSL**.
4. Dans le projet de complément, ouvrez le fichier XML de manifeste. Veillez à modifier le code XML source. Pour certains types de projets, Visual Studio ouvre un affichage visuel du code XML qui ne fonctionnera pas pour l’étape suivante.
5. Cherchez toutes les instances de **~remoteAppUrl/** et remplacez-les par l’URL SSL que vous venez de copier. Vous verrez plusieurs remplacements en fonction du type de projet, et les nouvelles URL ressembleront à `https://localhost:44300/Home.html`.
6. Enregistrez le fichier XML.
7. Cliquez avec le bouton droit sur le projet web, puis sélectionnez **Déboguer** -> **Démarrer une nouvelle instance**. Cela a pour effet d’exécuter le projet web sans lancer Office.
8. À partir d’Office sur le web, chargez la version test du complément en suivant les étapes décrites précédemment dans [Chargement de version test d’un complément Office dans Office sur le web](#sideload-an-office-add-in-in-office-on-the-web).

## <a name="remove-a-sideloaded-add-in"></a>Supprimer un complément versions test chargées

Vous pouvez supprimer un complément précédemment versions test chargées en effaçant le cache de votre navigateur. En outre, si vous modifiez le manifeste de votre complément (par exemple, mettez à jour les noms de fichier des icônes ou du texte de commandes de complément), vous devrez peut-être effacer le cache, puis rechargementer le complément à l’aide d’un manifeste mis à jour. Cette action permettra à Office d’afficher le complément tel que décrit par le manifeste mis à jour.
